#!/usr/bin/env python3
"""HTML 一頁摘要產生器（A5；改寫自 OpenAI v0.6 的 tools/build_html.py）：
LibreOffice 重算 → 以具名範圍或「頁!列 ID」取值 → 單一靜態 HTML。

- 計算一律由 Excel 執行：基準與敏感度情境都是「改寫 Inputs（或 TK NonNV 快照格）→ LibreOffice 重算」後讀值；
  本程式只做取值、格式化與排版（含 SVG 幾何），不做任何模型計算。
- 敏感度情境定義在 tools/html_scenarios.yaml（值取自 Inputs 的低／高欄、TK_Link 的 Lo／Hi 具名格）。
- 每個數字以 <span class="n" data-ref data-y data-sc data-fmt data-v> 標示來源，title 顯示 Excel 位置；
  tests/test_html.py 以 engine（pycel）逐一比對。
- 取位寫法：具名範圍（多年範圍需指定年度）；「頁!列 ID」（年度欄 D–I＝2025–2030，未指定年度＝D 欄）；
  「頁!列 ID:欄」（指定欄，例 Funding!F88:G＝回流對照表的「模型付款」欄、SRC_ANT!SRC_ANT_373:F＝低欄）。
- 輸出單一檔案、離線可開：CSS 內嵌、圖表為內嵌 SVG、系統字型，無任何外部 URL。

用法：python3 tools/build_html.py [--outdir dist]
"""
from __future__ import annotations

import argparse
import html
import json
import os
import math
import re
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
import yaml

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests" / "parity"))
from engine import current_model_path, parse_ref  # noqa: E402

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
YEAR_COL = {y: chr(ord("D") + i) for i, y in enumerate(YEARS)}   # D–I＝2025–2030（各頁共同版面）
SCEN_FILE = REPO / "tools" / "html_scenarios.yaml"
MINUS = "−"
ROWID_RE = re.compile(r"[A-Z]{1,3}\d{1,3}|C\d{2,3}|INP_\d{3}|SRC_ANT_\d{3}")


# ── 取位：具名範圍或「頁!列 ID[:欄]」→（頁, 格, 列 ID）──────────────────────
class Locator:
    """以公式模式的活頁簿建立位址對照（不讀計算值）。"""

    def __init__(self, path: Path):
        wb = openpyxl.load_workbook(path)
        self.wb = wb
        self.names = {k: v.attr_text for k, v in wb.defined_names.items()}
        self.rowid: dict[tuple[str, str], int] = {}       # (頁, 列 ID) → 列號
        self.rowid_of: dict[tuple[str, int], str] = {}    # (頁, 列號) → 列 ID
        for ws in wb:
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if isinstance(v, str) and ROWID_RE.fullmatch(v):
                    self.rowid.setdefault((ws.title, v), r)
                    self.rowid_of[(ws.title, r)] = v
        self.inputs = wb["Inputs"]

    def locate(self, ref: str, year: int | None = None) -> tuple[str, str, str]:
        if "!" in ref:
            sheet, rid = ref.split("!", 1)
            col = None
            if ":" in rid:
                rid, col = rid.split(":", 1)
            if re.fullmatch(r"[A-Z]+\d+", rid) and (sheet, rid) not in self.rowid:   # 直接格位（TK_Link!B4 等）
                return sheet, rid, rid
            r = self.rowid[(sheet, rid)]
            col = col or (YEAR_COL[year] if year is not None else "D")
            return sheet, f"{col}{r}", rid
        sheet, coord = parse_ref(self.names[ref])
        coord = coord.replace("$", "")
        if ":" in coord:
            a, _ = coord.split(":")
            r = int(re.sub(r"[A-Z]+", "", a))
            if year is None:
                raise ValueError(f"{ref} 是多年範圍，需指定年度")
            coord = f"{YEAR_COL[year]}{r}"
        r = int(re.sub(r"[A-Z]+", "", coord))
        return sheet, coord, self.rowid_of.get((sheet, r), "")

    def const(self, name: str):
        """單格具名範圍的常數值（公式模式活頁簿；只用於輸入格）。"""
        sheet, coord = parse_ref(self.names[name])
        return self.wb[sheet][coord.replace("$", "")].value

    # Inputs 的值／低／高欄（常數格）；TK NonNV 快照格的 Lo／Hi 具名格
    def input_cell(self, name: str, which: str):
        if name.startswith("TK_"):
            return self.const(name + {"val": "", "lo": "Lo", "hi": "Hi"}[which])
        sheet, coord = parse_ref(self.names[name])
        coord = coord.replace("$", "")
        r = int(re.sub(r"[A-Z]+", "", coord))
        col = {"val": "E", "lo": "F", "hi": "G"}[which]
        return self.inputs[f"{col}{r}"].value


def load_scenarios(path: Path = SCEN_FILE) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def resolve_overrides(scn_id: str, loc: Locator, cfg: dict | None = None) -> dict[str, float]:
    """情境 → {INP_nnn 或 TK_NNV_*: 值}；值一律取自 Excel（Inputs 低／高欄、TK_Link Lo／Hi）。"""
    cfg = cfg or load_scenarios()
    out: dict[str, float] = {}
    for g in cfg["scenarios"][scn_id]:
        for name, spec in cfg["groups"][g]["set"].items():
            if spec in ("lo", "hi"):
                v = loc.input_cell(name, spec)
            elif isinstance(spec, dict) and "mul" in spec:
                v = loc.input_cell(name, "val") * loc.input_cell(spec["mul"], "val")
            elif spec == "prev18":
                out.update(prev_cells(name, PREV18_XLSX))      # v0.1.9：TK 快照回到 v21.7（值取自 v0.1.8 Excel）
                continue
            elif spec == "prev16":
                out.update(prev_cells(name, PREV16_XLSX))      # v0.1.7：TK 快照回到 v21.2（值取自 v0.1.6 Excel）
                continue
            elif spec == "prev14":
                out.update(prev_cells(name, PREV14_XLSX))      # v0.1.5：TK 快照回到 v19.1（值取自 v0.1.4 Excel）
                continue
            elif spec == "prev19":
                out.update(prev_cells(name, PREV19_XLSX))      # v0.1.4：TK 快照回到 v19.0（值取自 v0.1.3 Excel）
                continue
            elif spec == "prev":
                out.update(prev_cells(name))       # v0.1.2：Tokenomics 快照回到上一版（v0.1.1＝TK v5.26）的值；值取自上一版 Excel
                continue
            elif isinstance(spec, (int, float)):
                v = spec
            else:
                raise ValueError(f"{g}.{name}: 不支援的寫法 {spec!r}")
            if v is None:
                raise ValueError(f"{g}.{name}: 欄位為空")
            out[name] = v
    return out


PREV_XLSX = REPO / "model" / "archive" / "20261021_Anthropic_v0.1.1.xlsx"
PREV19_XLSX = REPO / "model" / "archive" / "20261022_Anthropic_v0.1.3.xlsx"
PREV14_XLSX = REPO / "model" / "archive" / "20261022_Anthropic_v0.1.4.xlsx"
PREV16_XLSX = REPO / "model" / "archive" / "20261031_Anthropic_v0.1.6.xlsx"
PREV18_XLSX = REPO / "model" / "archive" / "20261103_Anthropic_v0.1.8.xlsx"
_PREV: dict = {}


def prev_cells(name: str, path: Path | None = None) -> dict[str, object]:
    """上一版（PREV_XLSX，v0.1.1：TK 快照 v5.26）TK_Link 具名範圍的值；單格 → {name: v}，多格 → {name[k]: v}（k 從 1 起）。"""
    path = path or PREV_XLSX
    if path not in _PREV:
        wb = openpyxl.load_workbook(path)
        _PREV[path] = (wb, {k: v.attr_text for k, v in wb.defined_names.items()})
    wb, names_ = _PREV[path]
    sheet, ref = parse_ref(names_[name])
    ws = wb[sheet]
    ref = ref.replace("$", "")
    if ":" not in ref:
        return {name: ws[ref].value}
    cells = [c.value for row in ws[ref] for c in row]
    return {f"{name}[{k + 1}]": v for k, v in enumerate(cells)}


# ── 格式化（測試共用）───────────────────────────────────────────────────
def fmt(v, kind: str) -> str:
    if kind == "yr":
        return str(int(round(v)))
    if kind == "yt":                       # 年度或文字（例：首次缺口年「無」）
        return str(int(round(v))) if isinstance(v, (int, float)) else str(v)
    if kind == "pct0":
        x = round(v * 100)
        return f"{MINUS if x < 0 else ''}{abs(x):d}%"
    if kind == "pct1":
        x = round(v * 100, 1)
        return f"{MINUS if x < 0 else ''}{abs(x):.1f}%"
    m = re.fullmatch(r"(s?)f(\d)", kind)  # sfN：正值加「＋」
    if m:
        n = int(m.group(2))
        x = round(v, n)
        if x == 0:
            x = 0.0
        sign = MINUS if x < 0 else ("＋" if (m.group(1) and x > 0) else "")
        return f"{sign}{abs(x):,.{n}f}"
    if kind == "text":
        return str(v)
    raise ValueError(kind)


# ── 值來源（LibreOffice 重算後的 xlsx）──────────────────────────────────
def _compact_values(path: Path) -> dict[str, dict[str, object]]:
    """LibreOffice 重算後 xlsx 的快取值 → {頁: {座標: 值}}（只讀值、不讀公式；空格不存）。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        d = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and hasattr(c, "coordinate"):
                    d[c.coordinate] = c.value
        out[ws.title] = d
    wb.close()
    return out


class Values:
    def __init__(self, loc: Locator, books: dict[str, Path]):
        self.loc = loc
        self.wb = {k: _compact_values(p) for k, p in books.items()}   # v0.1.9：只存值（情境 150 個時省記憶體；不做計算）
        self.used: list[tuple] = []
        self.default_sc = "base"     # v0.1.7：頁尾各版更新紀錄以 v016（v0.1.6 當時的輸出）為預設情境，內容與前版相同

    def get(self, ref: str, year: int | None = None, sc: str | None = None):
        sc = sc or self.default_sc
        sheet, coord, _ = self.loc.locate(ref, year)
        v = self.wb[sc][sheet].get(coord)
        if v is None or v == "":
            raise ValueError(f"{sc}:{ref}@{year} 為空（{sheet}!{coord}）")
        if isinstance(v, str) and v.startswith("#"):
            raise ValueError(f"{sc}:{ref}@{year} 錯誤值 {v}")
        return v

    def title(self, ref: str, year: int | None, sc: str) -> str:
        sheet, coord, rid = self.loc.locate(ref, year)
        parts = [f"{sheet} {rid}".strip() if rid else sheet, coord]
        if "!" not in ref:
            parts.append(ref)
        if year is not None:
            parts.append(str(year))
        if sc != "base":
            parts.append(f"情境 {sc}")
        return "｜".join(parts)

    def n(self, ref: str, year: int | None = None, f: str = "f1", sc: str | None = None, cls: str = "") -> str:
        """可追溯的數字 span。"""
        sc = sc or self.default_sc
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        c = f"n {cls}".strip()
        return (f'<span class="{c}" data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" '
                f'data-v="{html.escape(repr(v))}" title="{html.escape(self.title(ref, year, sc))}">{html.escape(fmt(v, f))}</span>')

    def t(self, ref: str, year: int | None = None, f: str = "f1", sc: str | None = None, **attrs) -> str:
        """SVG <text>，同樣帶 data-ref（測試一併比對）。"""
        sc = sc or self.default_sc
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        a = " ".join(f'{k.replace("_", "-")}="{val}"' for k, val in attrs.items())
        return (f'<text {a} data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" data-v="{html.escape(repr(v))}">'
                f'<title>{html.escape(self.title(ref, year, sc))}</title>{html.escape(fmt(v, f))}</text>')

    def src(self, *refs: str) -> str:
        """圖表下方小字：Excel 位置（頁＋列 ID，具名範圍附註）。"""
        out = []
        for r in refs:
            multi = "!" not in r and ":" in self.loc.names.get(r, "")
            sheet, _, rid = self.loc.locate(r, YEARS[0] if ("!" in r or multi) else None)
            out.append(f"{sheet} {rid}" + ("" if "!" in r else f"（{r}）"))
        return '<p class="src">Excel：' + "、".join(dict.fromkeys(out)) + "</p>"


# ── SVG 小工具 ─────────────────────────────────────────────────────────
def nice_max(x: float) -> float:
    """座標軸上限：取 {1, 1.2, 1.6, 2, 2.4, 4, 6, 8, 10} × 10^k 中第一個 ≥ x 的值（4 等分刻度皆為整齊值）。"""
    p = 10 ** math.floor(math.log10(x))
    for m in (1, 1.2, 1.6, 2, 2.4, 4, 6, 8, 10):
        if m * p >= x - 1e-12:
            return m * p
    return 10 * p


def nice_range(lo: float, hi: float, n: int = 5) -> tuple[float, float, int]:
    """含正負的座標軸：步長取 {1, 2, 2.5, 5} × 10^k，回傳（下限, 上限, 刻度數）。"""
    raw = (hi - lo) / n
    p = 10 ** math.floor(math.log10(raw))
    step = next(m * p for m in (1, 2, 2.5, 5, 10) if m * p >= raw - 1e-12)
    a, b = math.floor(lo / step) * step, math.ceil(hi / step) * step
    return a, b, int(round((b - a) / step))


def axis(W, H, L, R, T, B, ymin, ymax, ticks, unit=""):
    out = []
    for i in range(ticks + 1):
        val = ymin + (ymax - ymin) * i / ticks
        y = T + (H - T - B) * (1 - (val - ymin) / (ymax - ymin))
        out.append(f'<line class="grid{" zero" if abs(val) < 1e-9 else ""}" x1="{L}" x2="{W - R}" y1="{y:.1f}" y2="{y:.1f}"/>')
        lab = f"{MINUS if val < -1e-9 else ''}{abs(val):g}"
        out.append(f'<text class="tick" x="{L - 6}" y="{y + 4:.1f}" text-anchor="end">{lab}</text>')
    if unit:
        out.append(f'<text class="tick" x="2" y="{T - 10}">{unit}</text>')
    return out


def ymap(val, H, T, B, ymin, ymax):
    return T + (H - T - B) * (1 - (val - ymin) / (ymax - ymin))


# ── 頁面 ───────────────────────────────────────────────────────────────
CSS = r"""
:root{--bg:#fbfaf7;--fg:#1d2329;--muted:#5d6670;--line:#d9dde1;--card:#ffffff;--accent:#0d5c63;
--rev:#2563a8;--cost:#c2410c;--gap:#b42318;--pos:#2f855a;--sub:#2563a8;--seat:#7fb2e5;--api:#5aa0d8;--ads:#a3c9ea;--net:#1d2329;--tgt:#7a3e9d;
--inf:#2f855a;--rd:#9ae6b4;--con:#4a5568;--own:#a0aec0;--amz:#d69e2e;--oai:#7a3e9d;--warn-bg:#fff4e5;--ok:#2f855a}
@media (prefers-color-scheme: dark){:root{--bg:#14181c;--fg:#e8ebee;--muted:#9aa4ae;--line:#323a42;--card:#1b2127;--accent:#5fc4cb;
--rev:#6aa7ea;--cost:#f08a4b;--gap:#ff7b72;--pos:#56c288;--sub:#6aa7ea;--seat:#3f6f9f;--api:#3f7fbf;--ads:#9cc7ee;--net:#e8ebee;--tgt:#c792ea;
--inf:#56c288;--rd:#2e6b4a;--con:#a0aec0;--own:#5a6573;--amz:#e6b450;--oai:#c792ea;--warn-bg:#2b2416;--ok:#56c288}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC","Microsoft JhengHei","Noto Sans CJK TC","Noto Sans TC","Heiti TC",sans-serif;font-variant-numeric:tabular-nums}
main{max-width:980px;margin:0 auto;padding:24px 16px 48px}
header h1{font-size:22px;line-height:1.3;margin:0 0 4px}
header .meta{color:var(--muted);font-size:13px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px 18px 14px;margin:16px 0}
h2{font-size:17px;margin:0 0 10px;display:flex;gap:8px;align-items:baseline}
h2 .no{color:var(--accent);font-weight:700}
h3{font-size:15px;margin:14px 0 6px}
p{margin:6px 0}
.lead{font-size:16px;font-weight:600;border-left:4px solid var(--gap);padding:6px 0 6px 12px;margin:4px 0 14px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.kpi{border:1px solid var(--line);border-radius:8px;padding:10px 12px}
.kpi .k{font-size:12px;color:var(--muted)}
.kpi .v{font-size:24px;font-weight:700;line-height:1.25}
.kpi .s{font-size:12px;color:var(--muted)}
.neg{color:var(--gap)}
.pos{color:var(--pos)}
.n{white-space:nowrap}
.src{font-size:11.5px;color:var(--muted);margin:4px 0 0}
.note{font-size:13px;color:var(--muted)}
.cw{overflow-x:auto;-webkit-overflow-scrolling:touch}
.tw{overflow-x:auto;-webkit-overflow-scrolling:touch;margin:8px 0}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:5px 6px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}
th:first-child,td:first-child{text-align:left;white-space:normal;min-width:9em}
thead th{color:var(--muted);font-weight:600;border-bottom:1.5px solid var(--muted)}
tr.em td{font-weight:700}
tr.sep td{border-top:1.5px solid var(--muted)}
svg{width:100%;height:auto;display:block}
svg text{fill:var(--fg);font-size:12px;font-family:inherit}
svg .tick{fill:var(--muted);font-size:11px}
svg .grid{stroke:var(--line);stroke-width:1}
svg .grid.zero{stroke:var(--muted)}
svg .lbl{font-size:11.5px}
svg .gaplbl{fill:var(--gap);font-weight:700;font-size:12px}
svg .poslbl{fill:var(--pos);font-weight:700;font-size:12px}
.legend{display:flex;flex-wrap:wrap;gap:4px 14px;font-size:12.5px;color:var(--muted);margin:2px 0 4px}
.legend i{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:-1px}
.must{background:var(--warn-bg);border-radius:8px;padding:10px 12px;margin:10px 0}
ol,ul{padding-left:1.3em;margin:6px 0}
li{margin:4px 0}
.tag{display:inline-block;font-size:11.5px;border:1px solid var(--line);border-radius:4px;padding:0 5px;margin-right:4px;color:var(--muted)}
dl{display:grid;grid-template-columns:max-content 1fr;gap:4px 14px;font-size:13px;margin:6px 0}
dt{color:var(--muted)}dd{margin:0;word-break:break-all}
footer{color:var(--muted);font-size:12px;text-align:center;margin-top:20px}
@media (max-width:640px){
 main{padding:16px 12px 40px}
 .kpis{grid-template-columns:1fr 1fr}
 .kpi .v{font-size:20px}
 section{padding:14px 12px 10px}
 header h1{font-size:19px}
 table{font-size:11.5px}
 th,td{padding:4px 3px}
 th:first-child,td:first-child{min-width:6.5em}
 .cw svg{min-width:560px}
}
@media print{section{break-inside:avoid}}
"""


def chart_gw(V: Values) -> str:
    """每 VR 等值 GW 營收 vs 全成本。2025 分母極小（數值為其他年的 5–15 倍），軸以 2026–2030 定上限，2025 長條截斷並標實值。"""
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    rev = [V.get("COST_PropRev_VR", y) for y in YEARS]
    full = [V.get("COST_PropFull_VR", y) for y in YEARS]
    ymax = nice_max(max(full[1:] + rev[1:]) * 1.12)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="每 VR 等值 GW 營收與全成本 2025–2030">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B/GW/年")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.3
    for i, y in enumerate(YEARS):
        x0 = L + slot * i + slot * 0.18
        clipped = False
        for j, (vals, cls, ref) in enumerate(((rev, "rev", "COST_PropRev_VR"), (full, "cost", "COST_PropFull_VR"))):
            v = vals[i]
            yt = ymap(min(v, ymax), H, T, B, 0, ymax)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
            if v > ymax:
                clipped = True
                xb = x0 + j * bw
                o.append(f'<path d="M{xb - 1:.1f},{T + 14:.1f} l{bw:.1f},-6 v5 l-{bw:.1f},6 z" style="fill:var(--card)"/>')
                o.append(V.t(ref, y, "f0", x=f"{xb + bw / 2 - 1:.1f}", y=f"{T + 34 + 14 * j:.1f}", text_anchor="middle",
                             **{"class": "lbl"}, style="fill:#fff;font-weight:700"))
        top = ymap(min(max(full[i], rev[i]), ymax), H, T, B, 0, ymax)
        g = V.get("COST_PropGap_VR", y)
        o.append(V.t("COST_PropGap_VR", y, "sf1", x=f"{x0 + bw:.1f}", y=f"{(T - 4 if clipped else top - 6):.1f}", text_anchor="middle",
                     **{"class": "gaplbl" if g < 0 else "poslbl"}))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def chart_rev(V: Values) -> str:
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    gross = [V.get("REV_GrossCapped", y) for y in YEARS]
    ymax = nice_max(max(gross) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="營收結構（個人訂閱、企業席位、API、其他）與管理層目標">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.36
    for i, y in enumerate(YEARS):
        xc = L + slot * i + slot * 0.42
        base = 0.0
        for ref, cls in (("REV_SubConsumer", "sub"), ("REV_SubSeats", "seat"), ("REV_API", "api"), ("REV_Other", "ads")):
            v = V.get(ref, y)
            if v <= 0:
                continue
            y1, y0 = ymap(base + v, H, T, B, 0, ymax), ymap(base, H, T, B, 0, ymax)
            o.append(f'<rect x="{xc - bw / 2:.1f}" y="{y1:.1f}" width="{bw:.1f}" height="{max(y0 - y1, 0):.1f}" style="fill:var(--{cls})"/>')
            base += v
        net = V.get("REV_NetCapped", y)
        yn = ymap(net, H, T, B, 0, ymax)
        o.append(f'<line x1="{xc - bw / 2 - 4:.1f}" x2="{xc + bw / 2 + 4:.1f}" y1="{yn:.1f}" y2="{yn:.1f}" style="stroke:var(--net);stroke-width:2.5"/>')
        o.append(V.t("REV_GrossCapped", y, "f1", x=f"{xc:.1f}", y=f"{ymap(base, H, T, B, 0, ymax) - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        if 2026 <= y <= 2029:
            yt = ymap(V.get("RVS_Target", y), H, T, B, 0, ymax)
            xd = xc + bw / 2 + 8
            o.append(f'<path d="M{xd:.1f},{yt - 6:.1f} l6,6 l-6,6 l-6,-6 z" style="fill:var(--tgt)"/>')
            o.append(V.t("RVS_Target", y, "f0", x=f"{xd + 9:.1f}", y=f"{yt + 4:.1f}", **{"class": "lbl"}, style="fill:var(--tgt)"))
        o.append(f'<text class="tick" x="{xc:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


# run-rate 里程碑（D23；只對照）：SRC_ANT 列 ID（每月一筆；同月重複報導取一筆）。月份位置取自 SRC_ANT 的「期間」欄。
RUNRATE_IDS = ["SRC_ANT_008", "SRC_ANT_011", "SRC_ANT_012", "SRC_ANT_013", "SRC_ANT_015", "SRC_ANT_016", "SRC_ANT_017",
               "SRC_ANT_019", "SRC_ANT_020", "SRC_ANT_021", "SRC_ANT_022", "SRC_ANT_023"]
RUNRATE_RANGE_ID = "SRC_ANT_373"      # 2026 年底投資人預期（只有低、高）


def period_pos(loc: Locator, sid: str) -> tuple[float, str]:
    """SRC_ANT 期間欄（I）→ 時間軸位置（年＋月份比例）；月初＝月首、月底＝月末、其餘＝月中。"""
    r = loc.rowid[("SRC_ANT", sid)]
    txt = str(loc.wb["SRC_ANT"][f"I{r}"].value)
    m = re.match(r"(\d{4})-(\d{2})", txt)
    yr, mo = int(m.group(1)), int(m.group(2))
    off = 0.0 if "月初" in txt else (1.0 if "月底" in txt else 0.5)
    return yr + (mo - 1 + off) / 12, f"{yr}-{mo:02d}"


def chart_runrate(V: Values) -> str:
    W, H, L, R, T, B = 640, 280, 44, 10, 26, 34
    x0, x1 = 2024.75, 2027.0
    hi = V.get(f"SRC_ANT!{RUNRATE_RANGE_ID}:G")
    ymax = nice_max(hi * 1.08)
    X = lambda t: L + (W - L - R) * (t - x0) / (x1 - x0)  # noqa: E731
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="年化營收里程碑與模型年營收">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B（年化／年）")
    for yr in (2025, 2026, 2027):
        o.append(f'<line class="grid" x1="{X(yr):.1f}" x2="{X(yr):.1f}" y1="{T}" y2="{H - B}"/>')
    for yr in (2025, 2026):
        o.append(f'<text class="tick" x="{X(yr + 0.5):.1f}" y="{H - 12}" text-anchor="middle">{yr}</text>')
        v = V.get("REV_GrossCapped", yr)
        yy = ymap(v, H, T, B, 0, ymax)
        o.append(f'<rect x="{X(yr) + 2:.1f}" y="{yy:.1f}" width="{X(yr + 1) - X(yr) - 4:.1f}" height="{H - B - yy:.1f}" style="fill:var(--rev);opacity:.22"/>')
        o.append(f'<line x1="{X(yr) + 2:.1f}" x2="{X(yr + 1) - 2:.1f}" y1="{yy:.1f}" y2="{yy:.1f}" style="stroke:var(--rev);stroke-width:2.5"/>')
        o.append(V.t("REV_GrossCapped", yr, "f1", x=f"{X(yr + 0.25):.1f}", y=f"{yy - 6:.1f}", text_anchor="middle", **{"class": "lbl"}, style="fill:var(--rev);font-weight:700"))
    pts = []
    for sid in RUNRATE_IDS:
        t, _ = period_pos(V.loc, sid)
        pts.append((X(t), ymap(V.get(f"SRC_ANT!{sid}:E"), H, T, B, 0, ymax), sid))
    o.append('<polyline fill="none" style="stroke:var(--cost);stroke-width:1.5" points="' + " ".join(f"{x:.1f},{y:.1f}" for x, y, _ in pts) + '"/>')
    for k, (x, y, sid) in enumerate(pts):
        o.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" style="fill:var(--cost)"/>')
        if k >= 9:   # 2026-04 起標值（其餘數值小，見明細表）
            o.append(V.t(f"SRC_ANT!{sid}:E", None, "f0", x=f"{x + 7:.1f}", y=f"{y + 4:.1f}", text_anchor="start", **{"class": "lbl"}, style="fill:var(--cost)"))
    t, _ = period_pos(V.loc, RUNRATE_RANGE_ID)
    xr = X(min(t, x1 - 0.03))
    ylo, yhi = ymap(V.get(f"SRC_ANT!{RUNRATE_RANGE_ID}:F"), H, T, B, 0, ymax), ymap(hi, H, T, B, 0, ymax)
    o.append(f'<line x1="{xr:.1f}" x2="{xr:.1f}" y1="{ylo:.1f}" y2="{yhi:.1f}" style="stroke:var(--tgt);stroke-width:5;stroke-linecap:round"/>')
    o.append(V.t(f"SRC_ANT!{RUNRATE_RANGE_ID}:F", None, "f0", x=f"{xr - 7:.1f}", y=f"{ylo + 4:.1f}", text_anchor="end", **{"class": "lbl"}, style="fill:var(--tgt)"))
    o.append(V.t(f"SRC_ANT!{RUNRATE_RANGE_ID}:G", None, "f0", x=f"{xr - 7:.1f}", y=f"{yhi + 4:.1f}", text_anchor="end", **{"class": "lbl"}, style="fill:var(--tgt)"))
    o.append("</svg>")
    return "\n".join(o)


def chart_compute(V: Values) -> str:
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    sup = [V.get("CMP_SupplyGW", y) for y in YEARS]
    ymax = nice_max(max(sup) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="算力需求與供給 GW">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "GW")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.32
    for i, y in enumerate(YEARS):
        x0 = L + slot * i + slot * 0.16
        stacks = ((("CMP_InfGW_Eff", "inf"), ("CMP_RDGW", "rd"), "CMP_DemandGW"),
                  (("CMP_SupplyContractGW", "con"), ("CMP_Supply_Owned", "own"), "CMP_SupplyGW"))
        for j, (a, b, tot) in enumerate(stacks):
            base = 0.0
            x = x0 + j * (bw + 2)
            for ref, cls in (a, b):
                v = V.get(ref, y)
                y1, y0 = ymap(base + v, H, T, B, 0, ymax), ymap(base, H, T, B, 0, ymax)
                if y0 - y1 > 0.05:
                    o.append(f'<rect x="{x:.1f}" y="{y1:.1f}" width="{bw:.1f}" height="{y0 - y1:.1f}" style="fill:var(--{cls})"/>')
                base += v
            o.append(V.t(tot, y, "f1", x=f"{x + bw / 2:.1f}", y=f"{ymap(base, H, T, B, 0, ymax) - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def chart_cash(V: Values) -> str:
    """年底現金：基準 vs S7 合併不利情境，最低現金以短橫線標示。"""
    W, H, L, R, T, B = 640, 280, 44, 10, 26, 34
    yrs = YEARS[1:]
    series = (("FND_CashEnd", "rev"), ("Funding!F79", "cost"))
    ymax = nice_max(max(V.get(r, y) for r, _ in series for y in yrs) * 1.12)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="年底現金：基準與 S7 合併不利情境">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B")
    slot = (W - L - R) / len(yrs)
    bw = slot * 0.3
    for i, y in enumerate(yrs):
        x0 = L + slot * i + slot * 0.18
        for j, (ref, cls) in enumerate(series):
            v = V.get(ref, y)
            yt = ymap(v, H, T, B, 0, ymax)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
            o.append(V.t(ref, y, "f0", x=f"{x0 + j * bw + bw / 2 - 1:.1f}", y=f"{yt - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        ym = ymap(V.get("FND_MinCash", y), H, T, B, 0, ymax)
        o.append(f'<line x1="{x0 - 4:.1f}" x2="{x0 + 2 * bw + 2:.1f}" y1="{ym:.1f}" y2="{ym:.1f}" style="stroke:var(--gap);stroke-width:2;stroke-dasharray:4 3"/>')
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def chart_pair(V: Values) -> str:
    """與 OpenAI v0.6 並排：每 VR 等值 GW 差額 2026–2030（2025 分母效應，兩家皆略）。"""
    W, H, L, R, T, B = 640, 260, 44, 10, 26, 34
    yrs = YEARS[1:]
    series = (("COST_PropGap_VR", "rev", "Anthropic"), ("OAI_COST_PropGap_VR", "oai", "OpenAI"))
    vals = [V.get(r, y) for r, _, _ in series for y in yrs]
    ymin, ymax, ticks = nice_range(min(vals + [0]) * 1.05, max(vals + [0]) * 1.15)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="每 VR 等值 GW 差額：Anthropic 與 OpenAI">']
    o += axis(W, H, L, R, T, B, ymin, ymax, ticks, "$B/GW/年")
    slot = (W - L - R) / len(yrs)
    for ref, cls, _ in series:
        pts = [(L + slot * i + slot / 2, ymap(V.get(ref, y), H, T, B, ymin, ymax)) for i, y in enumerate(yrs)]
        o.append(f'<polyline fill="none" style="stroke:var(--{cls});stroke-width:2.5" points="' + " ".join(f"{x:.1f},{y:.1f}" for x, y in pts) + '"/>')
        for (x, yy), y in zip(pts, yrs):
            o.append(f'<circle cx="{x:.1f}" cy="{yy:.1f}" r="4" style="fill:var(--{cls})"/>')
            above = cls == "rev"
            o.append(V.t(ref, y, "sf1", x=f"{x:.1f}", y=f"{yy - 9 if above else yy + 17:.1f}", text_anchor="middle", **{"class": "lbl"}, style=f"fill:var(--{cls});font-weight:700"))
    for i, y in enumerate(yrs):
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def year_table(V: Values, rows: list[tuple], years=YEARS, head="$B") -> str:
    """rows：(標籤, ref, fmt, 樣式)。"""
    h = "".join(f"<th>{y}</th>" for y in years)
    o = [f'<div class="tw"><table><thead><tr><th>{head}</th>{h}</tr></thead><tbody>']
    for lab, ref, f, cls in rows:
        cells = []
        for y in years:
            try:
                cells.append(f"<td>{V.n(ref, y, f)}</td>")
            except ValueError:
                cells.append("<td>—</td>")
        o.append(f'<tr class="{cls}"><td>{lab}</td>{"".join(cells)}</tr>')
    o.append("</tbody></table></div>")
    return "\n".join(o)


def cell(V: Values, ref: str, year=None, f="f1", sc=None) -> str:
    try:
        return V.n(ref, year, f, sc)
    except ValueError:
        return "—"


def build(outdir: Path) -> tuple[Path, Path, Values]:
    from parity_lib import lo_recalc, set_inputs

    model = current_model_path()
    loc = Locator(model)
    cfg = load_scenarios()
    # 環境變數 ANT_HTML_CACHE（只供重複產生時省時間）：情境重算結果放該資料夾，已存在且比模型新者沿用；不設時用暫存資料夾並於結束時刪除
    cache = os.environ.get("ANT_HTML_CACHE")
    tmp = Path(cache) if cache else Path(tempfile.mkdtemp(prefix="a5_html_"))
    books = {}
    for sc in cfg["scenarios"]:
        done = tmp / f"out_{sc}" / model.name
        if cache and done.exists() and done.stat().st_mtime > model.stat().st_mtime and (tmp / f"in_{sc}" / "overrides.json").exists() \
                and json.loads((tmp / f"in_{sc}" / "overrides.json").read_text(encoding="utf-8")) == {k: repr(v) for k, v in resolve_overrides(sc, loc, cfg).items()}:
            books[sc] = done
            continue
        src = tmp / f"in_{sc}" / model.name
        shutil.rmtree(src.parent, ignore_errors=True)
        src.parent.mkdir(parents=True)
        ov = resolve_overrides(sc, loc, cfg)
        if ov:
            set_inputs(model, src, ov)
        else:
            shutil.copy(model, src)
        books[sc] = lo_recalc(src, tmp / f"out_{sc}")
        (src.parent / "overrides.json").write_text(json.dumps({k: repr(v) for k, v in ov.items()}, ensure_ascii=False), encoding="utf-8")
    V = Values(loc, books)
    if V.get("CHK_Errors", None, "base") != 0:
        raise SystemExit("基準 CHK_Errors ≠ 0，不產生成品")

    stamp = model.name.split("_")[0]
    ver = re.search(r"_v([\d.]+)\.xlsx$", model.name).group(1)
    html_path = outdir / f"{stamp}_Anthropic收支模型_v{ver.replace('.', '_')}.html"
    xlsx_path = outdir / f"{stamp}_Anthropic收支模型_v{ver.replace('.', '_')}.xlsx"
    page = render(V, model, ver, stamp)
    outdir.mkdir(parents=True, exist_ok=True)
    html_path.write_text(page, encoding="utf-8")
    shutil.copyfile(model, xlsx_path)
    if not cache:
        shutil.rmtree(tmp, ignore_errors=True)
    return html_path, xlsx_path, V


def src_meta(model: Path) -> tuple[str, int, int]:
    """SRC_ANT 最新文件日期、列數、來源引用 IPO 招股書（草稿）者的列數（只用於資料說明，不進任何數字）。"""
    ws = openpyxl.load_workbook(model, read_only=True)["SRC_ANT"]
    ds, n, pros = [], 0, 0
    for row in ws.iter_rows(min_row=5, values_only=True):
        if not (isinstance(row[0], str) and row[0].startswith("SRC_ANT_")):
            continue
        n += 1
        d = row[13]
        if isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", d.strip()):
            ds.append(d.strip())
        s = f"{row[11] or ''} {row[22] or ''}"
        if "招股書" in s or "prospectus" in s.lower() or "說明書" in s:
            pros += 1
    return (max(ds) if ds else "—"), n, pros


def section_v020(V: "Values", n, ver: str) -> str:
    """①之九：v0.2.0 更新了什麼（初步招股說明書差異、拆解：只動 B1／B2／C／D1／E、估值頁摘要與敏感度）。只取值與排版，不做模型計算。"""
    Y = 2030
    diff = [("9-30 後新簽或修訂（C1）", "SRC_ANT_506", "SRC_ANT_572", "f1"), ("承諾總額", "SRC_ANT_495", "SRC_ANT_567", "f1"),
            ("附條件或可終止（C3）", "SRC_ANT_508", "SRC_ANT_569", "f1"), ("Microsoft 合約下限（C2）", "SRC_ANT_514", "SRC_ANT_573", "f1"),
            ("未來 12 個月預計支付", "SRC_ANT_509", "SRC_ANT_570", "f1"), ("其中附條件", "SRC_ANT_510", "SRC_ANT_571", "f1"),
            ("Google 2026 NPA 剩餘（B2）", "SRC_ANT_482", "SRC_ANT_564", "f1"), ("AMD 股權承諾（B2）", "SRC_ANT_332", "SRC_ANT_565", "f1"),
            ("IPO 募資（報導 → 招股書淨募資 IPO＋私募，不含超額配售）", "SRC_ANT_341", "SRC_ANT_534", "f1")]
    diff_html = "".join(f"<tr><td>{lab}</td><td>{n(a, None, f)}</td><td>{n(b, None, f)}</td></tr>" for lab, a, b, f in diff)
    new = [("RSU 淨結算代扣稅款（以募資支付）", "SRC_ANT_537", "f1"), ("同步私募合計（Google 10、NVIDIA 10、AMD 2.5 等）", "SRC_ANT_525", "f1"),
           ("Run-Rate Revenue 2026-10-31", "SRC_ANT_557", "f1"), ("Run-Rate Revenue 2026-09-30（推算）", "SRC_ANT_559", "f1"),
           ("pro forma 現金與短投（9-30）", "SRC_ANT_560", "f3"), ("承諾中 2030-12-31 前發生（新揭露）", "SRC_ANT_568", "f1"),
           ("pro forma RSU 修改追加股權報酬（2026Q4）", "SRC_ANT_551", "f1"), ("未認列股權報酬（RSU；2.0 年）", "SRC_ANT_553", "f1"),
           ("Lambda 8 月合約（目前可終止；照計，C4）", "SRC_ANT_575", "f1"), ("發行後股數（百萬股）", "SRC_ANT_539", "f0")]
    new_html = "".join(f"<tr><td>{lab}</td><td>{n(a, None, f)}</td></tr>" for lab, a, f in new)
    cols = [("v0.1.9", "v019"), ("只動 B1 IPO", "v020_only_b1"), ("只動 B2 條件式", "v020_only_b2"), ("只動 C 承諾", "v020_only_c"), ("只動 D1 Q4 營收", "v020_only_d1"),
            ("只動 E 股權報酬", "v020_only_e"), (f"v{ver}（全部）", "base")]
    def g(lab, ref, yr, f, em=False, cs=cols):
        return f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cs) + "</tr>"
    rows = [g(f"營收總額 {y}（$B）", "REV_GrossCapped", y, "f2", y == 2026) for y in (2026, 2027, 2030)]
    rows += [g(f"營運自由現金流 {y}（Funding F07）", "FND_FCF", y, "sf2", y in (2029,)) for y in (2026, 2027, 2028, 2029, 2030)]
    rows += [g("營運自由現金流持續轉正年", "FND_FCFPosSustainYear", None, "yt"), g("累計外部資金需求 2026–2030", "FND_ExtNeedCum", Y, "f2", True),
             g("只靠已到位融資的現金谷底", "FND_CashNoExtTrough", None, "f2", True), g("2026 年底現金", "FND_CashEnd", 2026, "f2"),
             g("股權報酬 2026", "COST_SBC", 2026, "f2"), g("股權報酬 2027", "COST_SBC", 2027, "f2"),
             g("命題 1 差額 2026（$B/GW/年）", "COST_PropGap_VR", 2026, "sf2"), g("命題 1 差額 2030", "COST_PropGap_VR", Y, "sf2", True),
             g("覆蓋率 2030", "COST_Coverage", Y, "f3"), g("命題 1 持續轉正年", "COST_GapPosSustainYear", None, "yt"),
             g("合約付款 2028（$B）", "COST_ContractPay", 2028, "f2"), g("未歸屬餘額（$B）", "COST_S1Unalloc", None, "f2")]
    alt = [("基準", "base"), ("C4 扣除 Lambda 34.9", "lam_on"), ("超額配售計入", "oa_on"), ("只關 B1（IPO 不進基準）", "v020_no_b1"), ("只關 D1（Q4＝Q3）", "v020_no_d1"),
           ("只關 E（無 RSU 修改一次性、無下限）", "v020_no_e"),
           ("最低現金 3 個月", "mc_lo"), ("最低現金 12 個月", "mc_hi"), ("SpaceXAI S-1 路徑（R4）", "r4_on"), ("R4＋最低現金 3 個月", "r4_mc_lo"), ("R4＋最低現金 12 個月", "r4_mc_hi"),
           ("扣除附條件 186.2（R5d）", "r5d_on"), ("R4＋R5d", "r4_r5d"), ("並列：殘差法", "resid_on"), ("殘差法＋R4", "r4_resid"), ("捐贈只計 Q3 一次", "chr_once"),
           ("未歸屬餘額舊讀法 快 2027–2031", "s1u_fast"), ("舊讀法 慢 2028–2036", "s1u_slow"), ("舊讀法 慢 2028–2037", "s1u_slow37"), ("未起租起租年 2026", "lease_lo"),
           ("未起租起租年 2028", "lease_hi"), ("SpaceXAI 以上限 120.5 等比例放大月費", "spx_cap"), ("Google 額外 TPU 計入（中點）", "gaddl_on"), ("Broadcom 起始年 2026", "bcm_2026"),
           ("AMD＝S-1 下限 20", "amd_s1"), ("只關 W2（Fluidstack 舊法）", "only_w2_off"), ("只關 W4", "only_w4_off"), ("W4 兩者皆不用", "w4_neither"), ("W4 只動 Q3 口徑", "w4_q3only"),
           ("W4 只動股權報酬", "w4_sbconly"), ("承諾不進基準（W1、W2 皆關）", "commit_off"), ("說明列：只關 W1", "commit_off_w1only")]
    alt_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("FND_ExtNeedCum", Y, "f2", sc)}</td><td>{n("FND_CashNoExtTrough", None, "f2", sc)}</td>'
                       f'<td>{n("COST_PropGap_VR", 2026, "sf2", sc)}</td><td>{n("COST_PropGap_VR", Y, "sf2", sc)}</td><td>{n("COST_GapPosSustainYear", None, "yt", sc)}</td>'
                       f'<td>{n("VAL_TP_W", None, "f2", sc)}</td><td>{n("VAL_Rating", None, "text", sc)}</td></tr>' for lab, sc in alt)
    val = [("無風險利率 rf（10 年美債，11-06）", "VAL_Rf", "pct1"), ("β（同業 5 年中位數 1.37 → Blume）", "VAL_Beta", "f3"), ("ERP", "VAL_ERP", "pct1"),
           ("ke＝rf＋β × ERP", "VAL_Ke", "pct1"), ("稅後債務成本", "VAL_KdAT", "pct1"), ("WACC", "VAL_WACC", "pct1"), ("永續成長率 g", "VAL_g", "pct1"), ("RONIC", "VAL_RONIC", "pct1"),
           ("情境 B：明確期現值（2026Q4–2032）", "VAL_PVExp_B", "f1"), ("情境 B：終值現值", "VAL_PVTV_B", "f1"), ("情境 B：企業價值 EV", "VAL_EV_B", "f1"),
           ("情境 B：終值占 EV", "VAL_TVShare_B", "pct0"), ("淨現金（pro forma 現金 − 融資租賃現值）", "VAL_NetCash", "f1"), ("評價股數（十億股）", "VAL_Shares", "f3"),
           ("情境 B：股權價值", "VAL_EqV_B", "f1"), ("DCF 每股：情境 A（API 任務數高端）", "VAL_TP_DCF_A", "f2"), ("DCF 每股：情境 B（基準）", "VAL_TP_DCF_B", "f2"),
           ("DCF 每股：情境 C（低端＋Q4＝Q3）", "VAL_TP_DCF_C", "f2"), ("DCF 目標價（情境權重 A／B／C 見 Inputs INP_260–262）", "VAL_TP_DCF", "f2"),
           ("NTM 營收總額（情境 B）", "VAL_NTMRev", "f2"), ("同業 EV／NTM 營收中位數（倍）", "VAL_PeerMult", "f2"), ("同業 TTM FCF 利益率中位數", "VAL_PeerFCFm", "pct1"),
           ("調整係數（情境 B）", "VAL_AdjFactor", "f3"), ("Comparable：市場主流（未調整）", "VAL_TP_CompMain", "f2"), ("Comparable：調整後（進加權）", "VAL_TP_CompAdj", "f2"),
           ("加權目標價", "VAL_TP_W", "f2"), ("上檔空間（對 $74）", "VAL_Up_W", "pct1"), ("評等", "VAL_Rating", "text")]
    val_html = "".join(f'<tr class="{"em" if ref in ("VAL_TP_W", "VAL_Rating") else ""}"><td>{lab}</td><td>{n(ref, None, f)}</td><td><code>{ref}</code></td></tr>' for lab, ref, f in val)
    sheet, coord, _ = V.loc.locate("VAL_SensBase")
    import re as _re
    ref = V.loc.names["VAL_SensTbl"].split("!")[1].replace("$", "")
    r0, r1 = (int(_re.sub(r"[A-Z]+", "", x)) for x in ref.split(":"))
    sens_html = "".join(f'<tr><td>{n(f"Val!A{r}", None, "text")}</td><td>{n(f"Val!B{r}", None, "f2")}</td><td>{n(f"Val!C{r}", None, "f2")}</td></tr>' for r in range(r0, r1 + 1))
    return f"""<section id="s1i"><h2><span class="no">①之九</span>v{ver} 更新了什麼：初步招股說明書（2026-11-10）滾入＋估值頁與評等</h2>
<p class="lead"><b>兩個命題都沒有翻轉；新增投資建議「{n("VAL_Rating", None, "text")}」。</b>IPO＋同步私募淨額 {n("FND_IPONet", None, "f1")} 進 2026 基準（J1），只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f1", "v019")} → {n("FND_CashNoExtTrough", None, "f1")}；
2026Q4 營收改用 10-31 年化 {n("SRC_ANT_557", None, "f0")} ÷ {n("INP_164", None, "f0")}（J2；中性近似）使 2027 起營收同比例上升、營運自由現金流 {n("FND_FCFPosSustainYear", None, "yt")} 年轉正（持續轉正年 {n("FND_FCFPosSustainYear", None, "yt", "v019")} → {n("FND_FCFPosSustainYear", None, "yt")}）；
IPO 時 RSU 一次性 {n("SRC_ANT_551", None, "f1")}（J3）與攤銷下限只加在全成本（非現金；估值 EBIT 加回，因這批 RSU 已計入評價股數），2026 差額 {n("COST_PropGap_VR", 2026, "sf2", "v019")} → {n("COST_PropGap_VR", 2026, "sf2")}，命題 1 持續轉正年 {n("COST_GapPosSustainYear", None, "yt", "v019")} → {n("COST_GapPosSustainYear", None, "yt")}；累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "v019")} → {n("FND_ExtNeedCum", Y, "f0")}。
估值（J5–J8）：DCF（FCFF）三情境加權 ${n("VAL_TP_DCF", None, "f2")}、調整後 Comparable ${n("VAL_TP_CompAdj", None, "f2")}，<b>加權目標價 ${n("VAL_TP_W", None, "f2")}，對 IPO 區間中點上檔空間 {n("VAL_Up_W", None, "pct1")}</b>；市場主流（未調整）Comparable ${n("VAL_TP_CompMain", None, "f2")}。
M 線規格「AI 收支模型不加 DCF 腿」（2026-11-05）因公司掛牌而失效。DCF 情境 B 的終值占 EV {n("VAL_TVShare_B", None, "pct0")}（>80%，明確期 2027–2028 自由現金流為負）；情境 C 的 EBIT 到 2030 仍為負，DCF 每股取下限（淨現金每股推到目標時點）{n("VAL_TP_DCF_C", None, "f2")}；DCF 與調整後 Comparable 的下限統一為淨現金每股 {n("VAL_NetCashPS", None, "f2")}（評價日）。</p>
<p class="note">評等說明：{n("VAL_RatingNote", None, "text")}</p>
<h3>初步招股說明書與 S-1/A（或前版來源）的差異</h3>
<div class="tw"><table><thead><tr><th>項目（$B）</th><th>舊值</th><th>初步招股說明書</th></tr></thead><tbody>{diff_html}</tbody></table></div>
<h3>初步招股說明書新揭露（只列示或已進模型）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>值</th></tr></thead><tbody>{new_html}</tbody></table></div>
<h3>拆解：只動該項（各自從 v0.1.9 加入）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in cols)}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>
<p class="note">B1 只改股權流入與現金路徑（自由現金流不變）；B2 只動條件式情境（S2、S3、資金第 3 層），基準不變；C 承諾改讀使未歸屬餘額與付款小幅增加；D1 使 2026Q4 營收 {n("REV_Q4_2026", None, "f2", "v019")} → {n("REV_Q4_2026", None, "f2")}、2027 起營收按校準 token 量同比例上升（2027 營收總額 ÷ 2026Q4 年化＝{n("REV_D2Ratio", None, "f3")}，≥ 1，D2 不立警示）；E 為非現金，只動全成本與估值的 EBIT。</p>
<h3>替代讀法，在 v{ver} 基準下（每項只改一處；v0.1.9 當時的同一組讀法見 ①之八）</h3>
<div class="tw"><table><thead><tr><th>讀法</th><th>累計外部資金</th><th>現金谷底</th><th>2026 差額</th><th>2030 差額</th><th>命題 1 持續轉正年</th><th>加權目標價</th><th>評等</th></tr></thead><tbody>{alt_html}</tbody></table></div>
<h3>估值頁（Val）摘要</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>值</th><th>Excel 具名範圍</th></tr></thead><tbody>{val_html}</tbody></table></div>
<h3>敏感度（加權目標價；一次改一個輸入，LibreOffice 重算整本活頁簿）</h3>
<div class="tw"><table><thead><tr><th>案例</th><th>低值情況</th><th>高值情況</th></tr></thead><tbody>{sens_html}</tbody></table></div>
<p class="note">基準加權目標價 {n("VAL_SensBase", None, "f2")}。永續成長率 g 在 RONIC＝WACC（價值中性）下不影響終值（NOPAT ×(1 − g ÷ WACC) ÷ (WACC − g)＝NOPAT ÷ WACC），所以 g 2%／4% 兩欄與基準相同。EV／NTM 毛利讀法：同業毛利率或 EV／毛利倍數未提供，列資料缺口未計算。</p>
{V.src("FND_IPONet", "FND_EquityIn", "REV_Q4_2026", "COST_SBC_IPO", "COST_SBC_Floor", "COST_ProsBy2030Ratio", "VAL_TP_W", "VAL_TP_DCF", "VAL_TP_CompAdj", "VAL_TP_CompMain", "VAL_Rating")}
</section>"""


def render(V: Values, model: Path, ver: str, stamp: str) -> str:
    n = V.n
    Y = 2030
    date_disp = f"{stamp[:4]}-{stamp[4:6]}-{stamp[6:]}"
    tk_file, tk_ver, tk_sha, tk_date = (V.get(f"TK_Link!B{i}") for i in (3, 4, 5, 6))
    oai_file, oai_sha, oai_commit = V.get("OAI_File"), V.get("OAI_SHA256"), V.get("OAI_Commit")
    src_date, src_n, src_pros = src_meta(model)
    chk_ids = sorted((k[1] for k in V.loc.rowid if k[0] == "Checks"), key=lambda s: int(s[1:]))
    report = f"docs/reports/{stamp}_v{ver}.md"
    _chk = json.loads((REPO / "builder" / "id_registry.json").read_text(encoding="utf-8"))["CHK"]
    chk_of = lambda k: _chk[k] + ":C"  # noqa: E731

    S = []
    # ── 標頭 ──
    S.append(f"""<header><h1>Anthropic 收支模型 v{ver}｜一頁摘要</h1>
<div class="meta">{date_disp}｜Excel <code>{html.escape(model.name)}</code>｜Tokenomics {html.escape(str(tk_ver))}（<code>{html.escape(str(tk_sha))[:7]}</code>）｜單位 $B（十億美元），曆年制｜數字游標停留可見 Excel 位置；手機上圖表可左右滑動｜<b>v0.1.7 起正式公開 S-1（2026-10-26）套用於基準：2025 校準、2026 上半年營收、逐家承諾與承諾三段、季度營業損益（非算力成本）。v0.1.7.1：原先二手轉載的 46 列經 chat 端 2026-11-02 以瀏覽器核對 SEC 原文、全部一致，改 Verified；2026 已到位融資依原文更正為 110.4（Google Series G-2 約 10 與 Series H 分開，v0.1.7 少算）。例外：AMD 維持模型 GW 型（S-1 只給下限 >20，另列情境）；股權報酬用 S-1「調整後 − GAAP 營業損益」代理（＝股權報酬 2.185＋非現金捐贈 0.126）；6-30 現金只對照。v0.1.8：S-1/A 第 1 號修訂（2026-11-02）的 2026Q3 實績、承諾總額 895.4 與逐年付款表套用於基準（三段 871.4、未歸屬餘額付款期依 S-1/A 結構）；附條件 213.8 照計、SpaceXAI S-1 路徑並列。v0.1.9：Tokenomics 快照 v21.7 → v21.10（TPU v7、Trainium3 產出比改 FP8 對齊作基準：0.8 → 0.66、0.5 → 0.28），η 重校；只改「每 VR 等值 GW」的換算，合約付款、現金與外部資金需求不變。<span style="color:var(--gap)">v0.2.0：初步招股說明書（2026-11-10；價格區間 $70–78）滾入：IPO＋同步私募淨額進 2026 基準、條件式改讀、承諾 900.9、2026Q4 營收改用 10-31 年化 ÷ 季數（中性近似）、IPO 時 RSU 一次性股權報酬與攤銷下限；新增估值頁（DCF＋Comparable）與評等。</span>見 ①之九、①之八、①之七、①之六、①之五</b></div></header>""")

    # ── ① 命題與結論 ──
    # 只做取值與排序（找谷底年、列出會產生外部資金需求的單一驅動），不做模型計算
    cash = {y: V.get("FND_CashNoExt", y) for y in YEARS[1:]}
    ty = min(cash, key=cash.get)
    if V.get("FND_ExtNeedCum", Y) > 0:
        p2_txt = (f'按目前已到位融資與基準支出路徑，<b>2026–2030 出現最低現金缺口</b>：累計外部資金需求 ${n("FND_ExtNeedCum", Y, "f1")}B（首次 {n("FND_FirstGapYear", None, "yt")} 年）；'
                  f'只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f1")}（{ty} 年），低於當年最低現金 {n("FND_MinCash", ty, "f1")}。金額小，但 v0.1.6 時谷底有 {n("FND_CashNoExtTrough", None, "f1", "v016")}——第 2 個結論已由「不需要」翻成「需要少量」。')
    else:
        p2_txt = (f'按目前已到位融資（正式 S-1 原文 ${n("FND_Committed", 2026, "f1")}B；S-1/A：Q3 無新股權）加 IPO＋私募淨額 ${n("FND_IPONet", None, "f1")}B（v0.2.0 B1）與基準支出路徑、<b>最低現金＝次年非算力成本 {n("INP_165", None, "f0")} 個月（Assumed 3–12）的假設下，2026–2030 不出現最低現金缺口</b>（累計外部資金需求 ${n("FND_ExtNeedCum", Y, "f0")}B）：'
                  f'只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f1")}（{ty} 年；當年最低現金 {n("FND_MinCash", ty, "f1")}，餘裕 {n("FND_Headroom", ty, "f1")}）。v0.1.8 起最低現金改 12 個月也不需要（{n("FND_ExtNeedCum", Y, "f1", "mc_hi")}；v0.1.7.1 時為 {n("FND_ExtNeedCum", Y, "f2", "v0171_mc_hi")}）。'
                  f'v0.2.0 起 IPO＋私募淨額進基準，單一讀法不再產生外部資金需求：SpaceXAI 改走 S-1 路徑（爬升後每月 3.1、上限 120.5）→ 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "r4_on")}、谷底 {n("FND_CashNoExtTrough", None, "f1", "r4_on")}（v0.1.9 時 {n("FND_ExtNeedCum", Y, "f1", "v019_r4_on")}）；未歸屬餘額回到 2027–2031 快付 → {n("FND_ExtNeedCum", Y, "f1", "s1u_fast")}、谷底 {n("FND_CashNoExtTrough", None, "f1", "s1u_fast")}——第 2 個結論更穩，但兩項以上同時落在下緣仍會需要外部資金（見 ⑥）。')
    S.append(f"""<section id="s1"><h2><span class="no">①</span>命題與一句話結論</h2>
<p class="note">命題（與 OpenAI 同一句）：Anthropic 每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本？若不能，缺口要多少外部資金、由誰補？</p>
<p class="note"><b>v0.2.0 摘要（初步招股說明書 2026-11-10＋估值）</b>：IPO＋同步私募淨額 {n("FND_IPONet", None, "f1")}（淨募資 {n("SRC_ANT_534", None, "f1")} − RSU 代扣稅款 {n("SRC_ANT_537", None, "f1")}，以區間中點 ${n("VAL_Px", None, "f0")}）進 2026 基準股權流入，只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f1", "v019")} → {n("FND_CashNoExtTrough", None, "f1")}；2026Q4 營收改用 10-31 Run-Rate Revenue {n("SRC_ANT_557", None, "f0")} ÷ {n("INP_164", None, "f0")}＝{n("REV_Q4_2026", None, "f2")}（v0.1.9 Q4＝Q3 {n("REV_Q4_2026", None, "f2", "v019")}），2026 營收總額 {n("REV_GrossCapped", 2026, "f2", "v019")} → {n("REV_GrossCapped", 2026, "f2")}，2027 起按校準 token 量同比例變動；營運自由現金流 {n("FND_FCFPosSustainYear", None, "yt")} 年轉正（2029 年 {n("FND_FCF", 2029, "sf1", "v019")} → {n("FND_FCF", 2029, "sf1")}），持續轉正年 {n("FND_FCFPosSustainYear", None, "yt", "v019")} → {n("FND_FCFPosSustainYear", None, "yt")}；IPO 時 RSU 一次性股權報酬 {n("SRC_ANT_551", None, "f1")} 與 2027–2028 攤銷下限（非現金）使 2026 差額 {n("COST_PropGap_VR", 2026, "sf2", "v019")} → {n("COST_PropGap_VR", 2026, "sf2")}，命題 1 持續轉正年 {n("COST_GapPosSustainYear", None, "yt")} 不變、命題 2 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0")} 不變。<b>估值：加權目標價 ${n("VAL_TP_W", None, "f2")}（DCF {n("VAL_TP_DCF", None, "f2")}、調整後 Comparable {n("VAL_TP_CompAdj", None, "f2")} 各 {n("INP_259", None, "pct0")}），對 IPO 區間中點 ${n("VAL_Px", None, "f0")} 上檔空間 {n("VAL_Up_W", None, "pct1")}，評等「{n("VAL_Rating", None, "text")}」</b>（見 ①之九）。</p>
<p class="lead">在目前需求成長、價格彈性、初步招股說明書揭露的算力承諾（表內 {n("SRC_ANT_496", None, "f1")}＋未起租 {n("SRC_ANT_504", None, "f1")}＋9-30 後新簽 {n("SRC_ANT_572", None, "f1")}＝${n("COST_S1fCommit3", None, "f0")}B；扣除模型逐家合約後的未歸屬餘額 ${n("COST_S1Unalloc", None, "f0")}B 依 S-1/A 付款表結構付款）、2026 年已到位融資（${n("FND_Committed", 2026, "f0")}B）與 IPO＋私募淨額（${n("FND_IPONet", None, "f1")}B）的假設下：<b>2026–2029 年每 VR 等值 GW 營收都不能覆蓋全成本</b>（{n("COST_PropGap_VR", 2026, "sf2")}／{n("COST_PropGap_VR", 2027)}／{n("COST_PropGap_VR", 2028)}／{n("COST_PropGap_VR", 2029)} $B/GW/年；2026 年由 v0.1.7.1 的 {n("COST_PropGap_VR", 2026, "sf2", "v0171")} 轉負：S-1/A 的 Q3 股權報酬 1.070 與非現金捐贈 1.960 依 Q4＝Q3 慣例計入全成本（R5a；皆非現金，不影響現金與命題 2），加上未歸屬餘額依 S-1/A 付款表自 2026Q4 開始付款（R5c）；見 ①之七），<b>{n("COST_GapPosSustainYear", None, "yt")} 年才持續轉正</b>（{n("COST_PropGap_VR", Y, "sf1")}，覆蓋率 {n("COST_Coverage", Y, "pct0")}）；
{p2_txt}
營運自由現金流 2026 年為正（{n("FND_FCF", 2026, "sf1")}）、2027–2028 年為負（2029 年 {n("FND_FCF", 2029, "sf1")}），{n("FND_FCFPosSustainYear", None, "yt")} 年起持續為正；若不靠任何融資（第 1 層），2026–2030 累計需外部提供 ${n("FND_L1_ExtTotal", None, "f0")}B；已到位股權（第 2 層）補上全部缺口（第 2 層累計外部資金需求 ${n("FND_L2_ExtTotal", None, "f0")}B，見 ⑤）。
<b>此結論要成立必須</b>：2029 年後不再大量新增算力合約（基準只計已簽合約與 S-1/A 承諾）、SpaceXAI 不按 S-1 路徑付滿上限、未歸屬餘額付款不早於 S-1/A 付款表結構，且 API 需求成長、價格反應不同時落在區間下緣——v0.2.0 起 IPO 淨額進基準後，單一項落在下緣只使現金谷底下降、不產生外部資金需求，兩項以上同時落在下緣才會（見 ⑥）。</p>
<div class="kpis">
<div class="kpi"><div class="k">2030 每 VR 等值 GW：營收 vs 全成本</div><div class="v">{n("COST_PropRev_VR", Y)} <span class="note">vs</span> {n("COST_PropFull_VR", Y)}</div><div class="s">$B/GW/年；差額 <b class="pos">{n("COST_PropGap_VR", Y, "sf1")}</b></div></div>
<div class="kpi"><div class="k">覆蓋率（營收淨額 ÷ 全成本）</div><div class="v">{n("COST_Coverage", Y, "pct0")}</div><div class="s">2030；2026–2029 為 {n("COST_Coverage", 2026, "pct0")}／{n("COST_Coverage", 2027, "pct0")}／{n("COST_Coverage", 2028, "pct0")}／{n("COST_Coverage", 2029, "pct0")}；全成本含股權報酬、現金口徑</div></div>
<div class="kpi"><div class="k">累計外部資金需求 2026–2030</div><div class="v">{n("FND_ExtNeedCum", Y, "f0")}</div><div class="s">$B；OpenAI v0.6.4 同期 <b class="neg">{n("OAI_FND_ExtNeedCum", Y, "f1")}</b></div></div>
<div class="kpi"><div class="k">只靠已到位融資的現金谷底</div><div class="v">{n("FND_CashNoExtTrough", None, "f1")}</div><div class="s">$B（v0.1.9 為 {n("FND_CashNoExtTrough", None, "f1", "v019")}、v0.1.7.1 為 {n("FND_CashNoExtTrough", None, "f1", "v0171")}），{ty} 年；當年最低現金 {n("FND_MinCash", ty, "f1")}</div></div>
</div>
<div class="kpis">
<div class="kpi"><div class="k">投資建議（評等）</div><div class="v">{n("VAL_Rating", None, "text")}</div><div class="s">加權目標價上檔空間 {n("VAL_Up_W", None, "pct1")}；現價＝IPO 區間中點 ${n("VAL_Px", None, "f2")}（區間 {n("VAL_PxLo", None, "f0")}–{n("VAL_PxHi", None, "f0")}）</div></div>
<div class="kpi"><div class="k">加權目標價（12 個月）</div><div class="v">${n("VAL_TP_W", None, "f2")}</div><div class="s">DCF {n("VAL_TP_DCF", None, "f2")} × {n("INP_259", None, "pct0")}＋調整後 Comparable {n("VAL_TP_CompAdj", None, "f2")} ×（1 − 權重）</div></div>
<div class="kpi"><div class="k">市場主流 Comparable 目標價</div><div class="v">${n("VAL_TP_CompMain", None, "f2")}</div><div class="s">同業 EV／NTM 營收中位數 {n("VAL_PeerMult", None, "f2")} 倍（未調整）；上檔空間 {n("VAL_Up_CompMain", None, "pct1")}</div></div>
<div class="kpi"><div class="k">上檔空間（加權）</div><div class="v">{n("VAL_Up_W", None, "pct1")}</div><div class="s">DCF 單獨 {n("VAL_Up_DCF", None, "pct1")}；WACC {n("VAL_WACC", None, "pct1")}、終值占 EV {n("VAL_TVShare_B", None, "pct0")}（情境 B）</div></div>
</div>
<p class="note">單一驅動落在區間下緣時現金谷底大幅下降，但 v0.2.0 起 IPO＋私募淨額進基準後都不再出現外部資金需求：API 任務數成長取低 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "api_task_lo")}（累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "api_task_lo")}；v0.1.9 時 {n("FND_ExtNeedCum", Y, "f1", "v019_api_task_lo")}）；價格彈性取弱端 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "elast_weak")}；SpaceXAI 走 S-1 路徑 → {n("FND_CashNoExtTrough", None, "f1", "r4_on")}；未歸屬餘額快付 → {n("FND_CashNoExtTrough", None, "f1", "s1u_fast")}；每 GW 合約價 20 → {n("FND_CashNoExtTrough", None, "f1", "price_hi")}。兩項同時落在下緣（任務數低＋彈性弱）→ 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "combo1")}（首次 {n("FND_FirstGapYear", None, "yt", "combo1")}）。附條件或可終止的 186.2 若扣除（R5d 讀法），未歸屬餘額 {n("COST_S1Unalloc", None, "f0", "r5d_on")}、現金谷底 {n("FND_CashNoExtTrough", None, "f1", "r5d_on")}、命題 1 {n("COST_GapPosSustainYear", None, "yt", "r5d_on")} 年持續轉正。正式承諾若不進基準（v0.1.6 合約口徑：W1、W2 皆關），2030 差額 {n("COST_PropGap_VR", Y, "sf1", "commit_off")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "commit_off")}、現金谷底 {n("FND_CashNoExtTrough", None, "f1", "commit_off")}。用管理層較舊的營收目標、同一支出，反向累計外部資金需求 {n("RVS_ExtNeedCum", Y, "f1")}（只作對照）。</p>
</section>""")

    # ①之二～①之四（各版更新紀錄）移到頁尾（Andy 2026-10-31：① 之後直接接 ② 的圖，較直觀）；
    # 此處照常產生，記下起點，產生完後整段移到 ⑨ 之後。
    _hist_start = len(S)
    S.append(section_v020(V, n, ver))
    # ── ①之八（v0.2.0 起以 v0.1.9 當時的輸出呈現：基準＝v019〔v0.2.0 新開關全關〕，現行讀法＝v019_ 版本；內容與前版相同） ──
    ver_cur, n_orig = ver, n
    ver = "0.1.9"
    V.default_sc = "v019"
    _m8 = {"base": "v019", **{k: "v019_" + k for k in ['mc_lo', 'mc_hi', 'r4_on', 'r4_mc_lo', 'r4_mc_hi', 'r5d_on', 'r4_r5d', 'resid_on', 'r4_resid', 'chr_once',
                                                      's1u_fast', 's1u_slow', 's1u_slow37', 'lease_lo', 'lease_hi', 'spx_cap', 'gaddl_on', 'bcm_2026', 'amd_s1',
                                                      'only_w2_off', 'only_w4_off', 'w4_neither', 'w4_q3only', 'w4_sbconly', 'commit_off', 'commit_off_w1only']}}
    n = lambda ref, year=None, f="f1", sc=None, cls="": n_orig(ref, year, f, _m8.get(sc, sc), cls)  # noqa: E731
    # ── ①之八 v0.1.9 更新了什麼（頁尾紀錄區最前；現行基準）：TK v21.7 → v21.10 ──
    def g8(lab, ref, yr, f, cols, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cols) + "</tr>")
    def rows8(cols):
        rr = [g8("η（2025 校準）", "CMP_Eta2025", None, "f4", cols, True)]
        rr += [g8(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", cols, y == 2030) for y in (2026, 2027, 2028, 2029, 2030)]
        rr += [g8("覆蓋率 2030（倍）", "COST_Coverage", 2030, "f3", cols), g8("命題 1 持續轉正年", "COST_GapPosSustainYear", None, "yt", cols),
               g8("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", 2030, "f2", cols, True), g8("首次缺口年", "FND_FirstGapYear", None, "yt", cols),
               g8("只靠已到位融資的現金谷底 $B（2029）", "FND_CashNoExtTrough", None, "f2", cols, True), g8("2029 最低現金 $B", "FND_MinCash", 2029, "f2", cols),
               g8("供給 VR 等值 GW 2030", "CMP_Supply_VReq", 2030, "f3", cols), g8("有效推論 GW 2030", "CMP_InfGW_Eff", 2030, "f3", cols),
               g8("研發 GW 2030", "CMP_RDGW", 2030, "f3", cols), g8("合約付款 2028 $B（應不變）", "COST_ContractPay", 2028, "f2", cols)]
        return rr
    wf9_cols = [("v0.1.8（TK v21.7）", "v018"), ("＋(a) TPU v7 產出比", "wf9_a"), ("＋(b) Trainium3 產出比", "wf9_ab"), (f"＋(c) 其他 DIFF＝v{ver}", "base")]
    one9_cols = [("v0.1.8", "v018"), ("只動 (a) TPU v7", "wf9_a"), ("只動 (b) Trainium3", "wf9_b"), ("只動 (c) 其他 DIFF", "wf9_c"), (f"v{ver}（三項）", "base")]
    off9_cols = [(f"v{ver} 基準", "base"), ("只關 (a)：TPU v7 回 v21.7", "v019_no_tpu"), ("只關 (b)：Trainium3 回 v21.7", "v019_no_trn"), ("全關＝v0.1.8", "v018")]
    alt9 = [(f"基準（v{ver}）", "base"), ("最低現金 3 個月", "mc_lo"), ("最低現金 12 個月", "mc_hi"), ("SpaceXAI S-1 路徑（R4）", "r4_on"), ("R4＋最低現金 3 個月", "r4_mc_lo"),
            ("R4＋最低現金 12 個月", "r4_mc_hi"), ("扣除附條件 213.8（R5d）", "r5d_on"), ("R4＋R5d", "r4_r5d"), ("並列：殘差法", "resid_on"), ("殘差法＋R4", "r4_resid"),
            ("捐贈只計 Q3 一次", "chr_once"), ("未歸屬餘額舊讀法 快 2027–2031", "s1u_fast"), ("舊讀法 慢 2028–2036", "s1u_slow"), ("舊讀法 慢 2028–2037", "s1u_slow37"),
            ("未起租起租年 2026", "lease_lo"), ("未起租起租年 2028", "lease_hi"), ("SpaceXAI 以上限 120.5 等比例放大月費", "spx_cap"), ("Google 額外 TPU 計入（中點）", "gaddl_on"),
            ("Broadcom 起始年 2026", "bcm_2026"), ("AMD＝S-1 下限 20", "amd_s1"), ("只關 W2（Fluidstack 舊法）", "only_w2_off"), ("只關 W4", "only_w4_off"),
            ("W4 兩者皆不用", "w4_neither"), ("W4 只動 Q3 口徑", "w4_q3only"), ("W4 只動股權報酬", "w4_sbconly"), ("承諾不進基準（W1、W2 皆關）", "commit_off"),
            ("說明列：只關 W1", "commit_off_w1only")]
    alt9_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("FND_ExtNeedCum", 2030, "f2", sc)}</td><td>{n("FND_FirstGapYear", None, "yt", sc)}</td>'
                        f'<td>{n("FND_CashNoExtTrough", None, "f2", sc)}</td><td>{n("COST_PropGap_VR", 2026, "sf2", sc)}</td><td>{n("COST_PropGap_VR", 2030, "sf2", sc)}</td>'
                        f'<td>{n("COST_GapPosSustainYear", None, "yt", sc)}</td></tr>' for lab, sc in alt9)
    sen9_cols = [(f"v{ver} 基準（FP8 對齊）", "base"), ("(i) TK 低端＝v21.9 原生精度", "sens9_lo"), ("(ii) TK 高端", "sens9_hi"), ("(iii) v21.7 舊值（對照）", "wf9_c")]
    tk9 = [("TPU v7 產出比 基準", "TK_NNV_TPUv7_Out"), ("　低", "TK_NNV_TPUv7_OutLo"), ("　高", "TK_NNV_TPUv7_OutHi"),
           ("Trainium3 產出比 基準", "TK_NNV_Trn3_Out"), ("　低", "TK_NNV_Trn3_OutLo"), ("　高", "TK_NNV_Trn3_OutHi"),
           ("TPU v7 持有比 基準（不變）", "TK_NNV_TPUv7_Hold"), ("Trainium3 持有比 基準（不變）", "TK_NNV_Trn3_Hold")]
    tk9_html = "".join(f"<tr><td>{lab}</td><td>{n(nm_, None, 'f2', 'v018')}</td><td>{n(nm_, None, 'f2')}</td></tr>" for lab, nm_ in tk9)
    ev9_cols = [("v0.1.1（TK v5.26）", "v011"), ("v0.1.2（v18.0）", "v012"), ("v0.1.3（v19.0）", "v013"), ("v0.1.4（v19.1）", "v014"), ("v0.1.5–v0.1.6（v21.2）", "v016"),
                ("v0.1.7–v0.1.8（v21.7）", "v018"), (f"v{ver}（{tk_ver}）", "base")]
    ev9_html = "".join(f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, None, 'f3', sc)}</td>" for _l, sc in ev9_cols) + "</tr>"
                       for lab, ref, em in (("η（2025）", "ETA_Eta", True), ("f_P 計價口徑", "ETA_fP", False), ("f_U 利用率", "ETA_fU", False), ("f_W 工作負載組合", "ETA_fW", False),
                                            ("未解釋殘差（不計生產折減）", "ETA_Resid", False), ("未解釋殘差（計入生產折減）", "ETA_Resid2", True)))
    S.append(f"""<section id="s1h"><h2><span class="no">①之八</span>v{ver} 更新了什麼：Tokenomics v21.7 → {html.escape(str(tk_ver))}（TPU v7、Trainium3 產出比改 FP8 對齊）</h2>
<p class="lead"><b>兩個命題都沒有翻轉。</b>Tokenomics 連升三版（v21.8 高成本情境 IT 折舊 4 → 5 年；v21.9 TPU／Trainium3 產出比改原生精度；v21.10 改 FP8 對齊作基準、原生精度降為低端）。本模型讀取的 129 個值中 16 個變動，進計算的只有 TPU v7、Trainium3 的產出比（持有比不變）。產出比下降 → 同樣的合約 GW 換成的 VR 等值 GW 變少（2030 年 {n("CMP_Supply_VReq", 2030, "f2", "v018")} → {n("CMP_Supply_VReq", 2030, "f2")}），2025 推論支出校準的 η 由 {n("CMP_Eta2025", None, "f3", "v018")} 升到 {n("CMP_Eta2025", None, "f3")}（回到 Checks 合理範圍 0.5–2，C54 警示消失）。<b>合約付款、營收、現金與外部資金需求逐位不變</b>（累計外部資金需求 {n("FND_ExtNeedCum", 2030, "f0")}、現金谷底 {n("FND_CashNoExtTrough", None, "f2")}，2029 年）；每 VR 等值 GW 的營收與全成本同比例放大，所以差額的絕對值變大（2026–2030 {n("COST_PropGap_VR", 2026, "sf2", "v018")}／{n("COST_PropGap_VR", 2027, "sf2", "v018")}／{n("COST_PropGap_VR", 2028, "sf2", "v018")}／{n("COST_PropGap_VR", 2029, "sf2", "v018")}／{n("COST_PropGap_VR", 2030, "sf2", "v018")} → {n("COST_PropGap_VR", 2026, "sf2")}／{n("COST_PropGap_VR", 2027, "sf2")}／{n("COST_PropGap_VR", 2028, "sf2")}／{n("COST_PropGap_VR", 2029, "sf2")}／{n("COST_PropGap_VR", 2030, "sf2")}），覆蓋率 {n("COST_Coverage", 2030, "f3")} 與持續轉正年 {n("COST_GapPosSustainYear", None, "yt")} 不變。有效推論與研發 GW 的分攤改變：2030 年推論 {n("CMP_InfGW_Eff", 2030, "f2", "v018")} → {n("CMP_InfGW_Eff", 2030, "f2")}、研發 {n("CMP_RDGW", 2030, "f2", "v018")} → {n("CMP_RDGW", 2030, "f2")}。</p>
<h3>Tokenomics NonNV 比例（進計算者）</h3>
<div class="tw"><table><thead><tr><th>比例</th><th>v21.7（v0.1.8）</th><th>{html.escape(str(tk_ver))}（v{ver}）</th></tr></thead><tbody>{tk9_html}</tbody></table></div>
<p class="note">其他 10 個變動名稱（IF_HoldEcon、IF_FullCost_*、IF_TrainCost_*、IF_FullCostDefault_*）只有「高成本」欄變（v21.8 IT 折舊 4 → 5 年，−9% 至 −11%）：本模型只用在高成本敏感度列（Cost 供應商持有成本・高成本、雲端毛利率・高成本）與 η 診斷的 f_P 高欄，基準欄未變，不進命題輸出。TPU v6e、Trainium2 沿用既有公式（Hopper × TPU／Trainium 產出比），新比例一併套用。</p>
<h3>拆解：從 v0.1.8 逐項加入</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in wf9_cols)}</tr></thead><tbody>{"".join(rows8(wf9_cols))}</tbody></table></div>
<h3>只動該項（各自從 v0.1.8 加入）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in one9_cols)}</tr></thead><tbody>{"".join(rows8(one9_cols))}</tbody></table></div>
<h3>反向檢查：從 v{ver} 只關該項</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in off9_cols)}</tr></thead><tbody>{"".join(rows8(off9_cols))}</tbody></table></div>
<h3>敏感度（不進基準）：TPU v7、Trainium3 產出比取 TK 低端／高端／v21.7 舊值</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in sen9_cols)}</tr></thead><tbody>{"".join(rows8(sen9_cols))}</tbody></table></div>
<p class="note">產出比取低端（＝v21.9 原生精度 {n("TK_NNV_TPUv7_OutLo", None, "f2")}／{n("TK_NNV_Trn3_OutLo", None, "f2")}）時 η 升到 {n("CMP_Eta2025", None, "f3", "sens9_lo")}，2030 差額 {n("COST_PropGap_VR", 2030, "sf2", "sens9_lo")}；取高端 η {n("CMP_Eta2025", None, "f3", "sens9_hi")}、差額 {n("COST_PropGap_VR", 2030, "sf2", "sens9_hi")}。三種讀法下命題 1（{n("COST_GapPosSustainYear", None, "yt", "sens9_lo")}／{n("COST_GapPosSustainYear", None, "yt", "sens9_hi")} 年持續轉正）與命題 2（累計外部資金需求 {n("FND_ExtNeedCum", 2030, "f0", "sens9_lo")}／{n("FND_ExtNeedCum", 2030, "f0", "sens9_hi")}）都不翻轉：產出比只決定「每 GW 能產多少 token」與 VR 等值換算，付款與營收由合約與需求決定，2025 η 校準吸收產能水準。</p>
<h3>既有預設的替代讀法，在 v{ver} 基準下（每項只改一處；v0.1.8 當時的同一組讀法見 ①之七）</h3>
<div class="tw"><table><thead><tr><th>讀法</th><th>累計外部資金</th><th>首次缺口年</th><th>現金谷底</th><th>2026 差額</th><th>2030 差額</th><th>命題 1 持續轉正年</th></tr></thead><tbody>{alt9_html}</tbody></table></div>
<p class="note">各讀法的現金欄（累計外部資金、首次缺口年、現金谷底）與 v0.1.8 相同；差額欄因 VR 等值 GW 變少而絕對值放大，正負號與持續轉正年不變。</p>
<h3>η 與 η 拆解：各 Tokenomics 版本</h3>
<div class="tw"><table><thead><tr><th>項目</th>{"".join(f"<th>{l}</th>" for l, _s in ev9_cols)}</tr></thead><tbody>{ev9_html}</tbody></table></div>
<p class="note">各欄為該版當時的輸出（歷史情境把 TK 快照與各版開關回到當時）。v21.10 的 η 變動全部落在未解釋殘差（f_P、f_U、f_W 不變）。OpenAI 並排仍為 v0.6.4（其 Tokenomics 為 v21.7），TPU／Trainium 口徑與本版不同，只影響 ⑦ 的並排列。</p>
{V.src("CMP_Eta2025", "CMP_Supply_VReq", "COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExtTrough", "CMP_InfGW_Eff", "CMP_RDGW", "COST_ContractPay", "ETA_Eta", "ETA_Resid2")}
</section>""")
    n, ver = n_orig, ver_cur
    # ── ①之七 v0.1.8 更新了什麼（v0.1.9 起以 v0.1.8 當時的輸出呈現） ──
    # v0.1.9 起以 v0.1.8 當時的輸出呈現：基準＝v018（TK 回到 v21.7），現行讀法的情境＝v018_ 版本；內容與前版相同
    ver7 = "0.1.8"
    V.default_sc = "v018"
    _m7 = {"base": "v018", **{k: "v018_" + k for k in ['r4_on', 'r5d_on', 'r4_r5d', 'mc_lo', 'mc_hi', 'r4_mc_lo', 'r4_mc_hi', 's1u_fast', 's1u_slow', 's1u_slow37', 'lease_lo', 'lease_hi',
                                                      'resid_on', 'r4_resid', 'spx_cap', 'gaddl_on', 'bcm_2026', 'amd_s1', 'only_w2_off', 'only_w4_off', 'w4_neither', 'w4_q3only',
                                                      'w4_sbconly', 'commit_off', 'commit_off_w1only', 'chr_once']}}
    n_cur7 = n
    n = lambda ref, year=None, f="f1", sc=None, cls="": n_cur7(ref, year, f, _m7.get(sc, sc), cls)  # noqa: E731
    def _r7(sheet, prefix):
        ws = V.loc.wb[sheet]
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and b.startswith(prefix):
                return f"{sheet}!{ws.cell(r, 1).value}"
        raise KeyError(f"{sheet} 找不到列：{prefix}")
    def g7(lab, ref, yr, f, cols, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cols) + "</tr>")
    def outrows(cols):
        rr = [g7(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", cols, y == 2026) for y in (2026, 2027, 2028, 2029, 2030)]
        rr += [g7("覆蓋率 2026（倍）", "COST_Coverage", 2026, "f3", cols), g7("覆蓋率 2030（倍）", "COST_Coverage", 2030, "f3", cols),
               g7("命題 1 持續轉正年", "COST_GapPosSustainYear", None, "yt", cols),
               g7("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", 2030, "f2", cols, True), g7("首次缺口年", "FND_FirstGapYear", None, "yt", cols),
               g7("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2", cols, True), g7("2029 最低現金 $B", "FND_MinCash", 2029, "f2", cols),
               g7("營運自由現金流持續轉正年", "FND_FCFPosSustainYear", None, "yt", cols)]
        return rr
    wf7_cols = [("v0.1.7.1", "v0171"), ("＋R5a Q3 實績", "wf8_a"), ("＋R5b 承諾 871.4", "wf8_ab"), ("＋R5c 付款期＝v0.1.8", "base")]
    wf7_rows = [g7("2026 營收總額 $B", "REV_Gross", 2026, "f2", wf7_cols), g7("2026 非算力成本（不含股權報酬）$B", "COST_NonCompExSBC", 2026, "f2", wf7_cols),
                g7("2026 股權報酬（含非現金捐贈）$B", "COST_SBC", 2026, "f2", wf7_cols), g7("S-1 承諾三段合計 $B", "COST_S1fCommit3", None, "f1", wf7_cols),
                g7("未歸屬餘額 $B", "COST_S1Unalloc", None, "f1", wf7_cols), g7("未歸屬餘額付款 2026 $B", "COST_Pay_S1Unalloc", 2026, "f2", wf7_cols),
                g7("未歸屬餘額付款 2028 $B", "COST_Pay_S1Unalloc", 2028, "f2", wf7_cols), g7("算力成本 2028 $B", "COST_Compute", 2028, "f1", wf7_cols)]
    wf7_rows += outrows(wf7_cols)
    one_cols = [("v0.1.7.1", "v0171"), ("只動 R5a", "wf8_a"), ("只動 R5b", "wf8_b"), ("只動 R5c", "wf8_c"), (f"v{ver7}（三項）", "base"),
                ("v0.1.8＋R4", "r4_on"), ("v0.1.8＋R5d", "r5d_on")]
    one_rows = outrows(one_cols)
    off_cols = [(f"v{ver7} 基準", "base"), ("只關 R5a", "v018_no_r5a"), ("只關 R5b", "v018_no_r5b"), ("只關 R5c", "wf8_ab"), ("全關＝v0.1.7.1", "v0171")]
    off_rows = outrows(off_cols)
    mc7 = [("3 個月（區間下緣）", "v0171_mc_lo", "mc_lo", "r4_mc_lo"), ("6 個月（基準）", "v0171", "base", "r4_on"), ("12 個月（區間上緣）", "v0171_mc_hi", "mc_hi", "r4_mc_hi")]
    mc7_html = "".join(
        f'<tr class="{"em" if b == "base" else ""}"><td>{lab}</td><td>{n("FND_MinCash", 2029, "f2", b)}</td>'
        f'<td>{n("FND_ExtNeedCum", 2030, "f2", a)}</td><td>{n("FND_FirstGapYear", None, "yt", a)}</td>'
        f'<td>{n("FND_ExtNeedCum", 2030, "f2", b)}</td><td>{n("FND_FirstGapYear", None, "yt", b)}</td><td>{n("FND_CashNoExtTrough", None, "f2", b)}</td>'
        + (f'<td>{n("FND_ExtNeedCum", 2030, "f2", c)}</td><td>{n("FND_FirstGapYear", None, "yt", c)}</td>' if c else "<td>—</td><td>—</td>") + "</tr>"
        for lab, a, b, c in mc7)
    alt7 = [(f"基準（v{ver7}：S-1/A 結構，未起租與 9-30 後 2027 起 10 年）", "base"), ("舊讀法：2027–2033 平均（R5c 關閉）", "wf8_ab"),
            ("舊讀法 快：2027–2031 平均", "s1u_fast"), ("舊讀法 慢：2028–2036 平均", "s1u_slow"), ("舊讀法 慢：2028–2037 平均", "s1u_slow37"),
            ("S-1/A 結構，未起租起租年 2026", "lease_lo"), ("S-1/A 結構，未起租起租年 2028", "lease_hi"), ("扣除附條件 213.8（R5d）", "r5d_on"),
            ("SpaceXAI S-1 路徑（R4）", "r4_on"), ("R4＋R5d", "r4_r5d"), ("並列：殘差法（S-1/A 結構 − 逐家合約時程）", "resid_on"), ("殘差法＋R4", "r4_resid")]
    alt7_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("COST_Pay_S1Unalloc", 2027, "f1", sc)}</td><td>{n("COST_Pay_S1Unalloc", 2028, "f1", sc)}</td>'
                        f'<td>{n("FND_ExtNeedCum", 2030, "f2", sc)}</td><td>{n("FND_FirstGapYear", None, "yt", sc)}</td><td>{n("FND_CashNoExtTrough", None, "f2", sc)}</td>'
                        f'<td>{n("COST_PropGap_VR", 2030, "sf2", sc)}</td><td>{n("COST_GapPosSustainYear", None, "yt", sc)}</td></tr>' for lab, sc in alt7)
    alt8 = [(f"基準（v{ver7}）", "base"), ("SpaceXAI 以上限 120.5 等比例放大月費（v0.1.7 W1(a)）", "spx_cap"), ("Google 額外 TPU 125–150 計入（中點）", "gaddl_on"),
            ("Broadcom 起始年改 S-1 期間起點 2026", "bcm_2026"), ("AMD 改讀正式 S-1 下限 20", "amd_s1"), ("只關閉 W2（自建回到 Fluidstack $50B 三年均攤）", "only_w2_off"),
            ("只關閉 W4（非算力成本回到 v0.1.6 算法）", "only_w4_off"), ("W4 兩者皆不用（下半年 Q2 持平、股權報酬模型值）", "w4_neither"), ("W4 只動 Q3 口徑（股權報酬回模型值）", "w4_q3only"),
            ("W4 只動股權報酬（下半年回 Q2 持平）", "w4_sbconly"), ("承諾不進基準（W1、W2 皆關）", "commit_off"), ("說明列：只關 W1", "commit_off_w1only")]
    alt8_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("FND_ExtNeedCum", 2030, "f2", sc)}</td><td>{n("FND_FirstGapYear", None, "yt", sc)}</td>'
                        f'<td>{n("FND_CashNoExtTrough", None, "f2", sc)}</td><td>{n("COST_PropGap_VR", 2026, "sf2", sc)}</td><td>{n("COST_PropGap_VR", 2030, "sf2", sc)}</td></tr>' for lab, sc in alt8)
    shr = [("S-1/A 表內義務逐年", _r7("Cost", "R5c S-1/A 表內義務逐年"), "f2"), ("未起租＋9-30 後新簽逐年", _r7("Cost", "R5c 未起租＋9-30 後新簽逐年"), "f2"),
           ("合成年度比例（v0.1.8）", "COST_S1aUnallocShare", "pct1"), ("舊：2027–2033 平均", _r7("Cost", "R5c 對照：v0.1.7 平均攤提"), "pct1"),
           ("未歸屬餘額付款（v0.1.8）", "COST_Pay_S1Unalloc", "f2")]
    shr_html = "".join(f"<tr><td>{lab}</td>" + "".join(f"<td>{n(ref, y, f)}</td>" for y in YEARS) + "</tr>" for lab, ref, f in shr)
    q3 = [("營收", "SRC_ANT_488", "SRC_ANT_384"), ("算力與基礎設施費用", "SRC_ANT_489", _r7("Cost", "R5a 對照：模型 Q3 隱含算力")),
          ("營業損益", "SRC_ANT_490", None), ("股權報酬", "SRC_ANT_491", None), ("非現金捐贈配對", "SRC_ANT_492", None), ("調整後營業損益", "SRC_ANT_493", None),
          ("非算力費用（不含股權報酬與捐贈）", "COST_S1aQ3NonComp", _r7("Cost", "R5a 對照：v0.1.7.1 的 Q3 隱含非算力")),
          ("股權報酬代理每季（股權報酬＋捐贈）", _r7("Cost", "R5a 股權報酬代理每季：Q3"), _r7("Cost", "R5a 對照：v0.1.7.1 每季代理"))]
    q3_html = "".join(f"<tr><td>{lab}</td><td>{n(a, 2026 if '!' in a else None, 'f3')}</td><td>{n(b, 2026 if '!' in b else None, 'f3') if b else '—'}</td></tr>"
                      for lab, a, b in q3)
    cm = [("表內義務（營運租賃＋融資租賃＋採購承諾）", "SRC_ANT_496", "進基準"), ("已簽未起租租賃（2026–2028 起租，加權約 10 年）", "SRC_ANT_504", "進基準"),
          ("9-30 後新簽或修訂", "SRC_ANT_506", "進基準"), ("短期承諾（一年內可無罰取消）", "SRC_ANT_503", "不進（N2）"), ("9-30 後短期", "SRC_ANT_507", "不進（N2）"),
          ("承諾總額（原文）", "SRC_ANT_495", "—"), ("其中附條件或可終止", "SRC_ANT_508", "基準計入；R5d 扣除讀法並列"),
          ("未來 12 個月預計支付（其中附條件 4.3）", "SRC_ANT_509", "只對照")]
    cm_html = "".join(f"<tr><td>{lab}</td><td>{n(r, None, 'f1')}</td><td>{t}</td></tr>" for lab, r, t in cm)
    spx = [("S-1 路徑年付款（R4）", "COST_S1aSpaceXPath"), ("現行基準 SpaceX 付款（月費 1.25 × 36）", "COST_Pay_Spacex")]
    spx_html = "".join(f"<tr><td>{lab}</td>" + "".join(f"<td>{n(ref, y, 'f2')}</td>" for y in YEARS) + "</tr>" for lab, ref in spx)
    c930 = [("S-1/A 9-30 現金及約當現金＋短期投資", "SRC_ANT_516", None, "f2"), ("模型近似（2025 年底＋已到位＋2026 自由現金流 × 9/12）", _r7("Funding", "模型 2026-09-30 近似"), 2026, "f2"),
            ("差（S-1/A − 模型）", "FND_Cash0930Gap", None, "sf2"), ("①股權收入：S-1/A 前九月特別股收入 − 模型已到位", _r7("Funding", "差額拆解 ①股權收入：S-1/A"), 2026, "sf3"),
            ("②前九月自由現金流時點（含租賃資產預付 1.955）", _r7("Funding", "差額拆解 ②前九月"), 2026, "sf2"), ("③受限現金 2.6", _r7("Funding", "差額拆解 ③受限現金：S-1/A"), 2026, "sf2"),
            ("④其餘", "FND_Cash0930Resid", None, "sf3")]
    c930_html = "".join(f"<tr><td>{lab}</td><td>{n(ref, yr, f)}</td></tr>" for lab, ref, yr, f in c930)
    S.append(f"""<section id="s1g"><h2><span class="no">①之七</span>v{ver7} 更新了什麼：S-1/A 第 1 號修訂（2026-11-02；Q3 實績、承諾總額 895.4 與逐年付款表）</h2>
<p class="lead"><b>命題 2 不翻轉、餘裕變厚；命題 1 的 2026 單年由正轉負，持續轉正年不變。</b>累計外部資金需求 2026–2030 維持 {n("FND_ExtNeedCum", 2030, "f0")}（v0.1.7.1 亦 {n("FND_ExtNeedCum", 2030, "f0", "v0171")}），只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f2", "v0171")} → <b>{n("FND_CashNoExtTrough", None, "f2")}</b>（2029 年；當年最低現金 {n("FND_MinCash", 2029, "f2")}），最低現金改 12 個月也不需要外部資金（v0.1.7.1 時需 {n("FND_ExtNeedCum", 2030, "f2", "v0171_mc_hi")}）。每 VR 等值 GW 差額 2026–2030 {n("COST_PropGap_VR", 2026, "sf2")}／{n("COST_PropGap_VR", 2027, "sf2")}／{n("COST_PropGap_VR", 2028, "sf2")}／{n("COST_PropGap_VR", 2029, "sf2")}／{n("COST_PropGap_VR", 2030, "sf2")}（v0.1.7.1：{n("COST_PropGap_VR", 2026, "sf2", "v0171")}／{n("COST_PropGap_VR", 2027, "sf2", "v0171")}／{n("COST_PropGap_VR", 2028, "sf2", "v0171")}／{n("COST_PropGap_VR", 2029, "sf2", "v0171")}／{n("COST_PropGap_VR", 2030, "sf2", "v0171")}），2030 覆蓋率 {n("COST_Coverage", 2030, "f3", "v0171")} → {n("COST_Coverage", 2030, "f3")}，持續轉正年 {n("COST_GapPosSustainYear", None, "yt")}。三項方向不同：R5a（Q3 實績）只讓 2026 全成本上升（股權報酬＋非現金捐贈每季 {n(_r7("Cost", "R5a 股權報酬代理每季：Q3"), 2026, "f2")}，皆非現金；一次性捐贈 {n("COST_S1aCharityInCost", 2026, "f2")} 只計 2026、不進每人基準），2027 起股權報酬與非算力現金成本都比 v0.1.7.1 低（2030 股權報酬 {n("COST_SBC", 2030, "f2", "v0171")} → {n("COST_SBC", 2030, "f2")}）；R5b（三段 832 → {n("COST_S1fCommit3", None, "f1")}、Google 114.5 → 121.3）單獨會使現金谷底降到 {n("FND_CashNoExtTrough", None, "f2", "wf8_b")}、需要 {n("FND_ExtNeedCum", 2030, "f2", "wf8_b")}；R5c（付款期改依 S-1/A 付款表：表內義務 2030 年後仍有 148.4、未起租 451.2 自 2027 起 10 年）把未歸屬餘額在 2026–2030 的付款比例由 {n(_r7("Cost", "R5c 對照：v0.1.7 平均攤提"), 2028, "pct1")}／年降到約 {n("COST_S1aUnallocShare", 2028, "pct1")}／年，抵銷 R5b 而有餘。<b>條件</b>：SpaceXAI 若照 S-1 路徑付滿 120.5（R4）→ 累計外部資金需求 {n("FND_ExtNeedCum", 2030, "f2", "r4_on")}（{n("FND_FirstGapYear", None, "yt", "r4_on")} 年），命題 2 翻為「需要」；附條件 213.8 若扣除（R5d）→ 谷底 {n("FND_CashNoExtTrough", None, "f2", "r5d_on")}、命題 1 {n("COST_GapPosSustainYear", None, "yt", "r5d_on")} 年持續轉正。</p>
<h3>瀑布：從 v0.1.7.1 逐項加入（R4、R5d 基準不動，只並列）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in wf7_cols)}</tr></thead><tbody>{"".join(wf7_rows)}</tbody></table></div>
<h3>只動該項的前後值（R5a、R5b、R5c 各自從 v0.1.7.1 加入；R4、R5d 從 v{ver7} 基準加入）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in one_cols)}</tr></thead><tbody>{"".join(one_rows)}</tbody></table></div>
<h3>反向檢查：從 v{ver7} 基準只關該項</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in off_cols)}</tr></thead><tbody>{"".join(off_rows)}</tbody></table></div>
<p class="note">開關（Inputs）：R4 s1f_spacex_fee_switch（基準 0）、R5a s1a_q3_switch、R5b s1a_commit_switch、R5c s1a_unalloc_switch（基準 1）、R5d s1a_cond_deduct_switch（基準 0）。R5a、R5b、R5c 全設 0 時，v0.1.7.1 的全部計算具名範圍逐位重現（engine 回歸測試）；逐格比對方法與格數見報告。</p>
<h3>最低現金月數 3／6／12（D17：基準 6 個月，Andy 未決；Assumed 3–12）</h3>
<div class="tw"><table><thead><tr><th>最低現金</th><th>2029 最低現金</th><th>v0.1.7.1 累計外部資金</th><th>v0.1.7.1 首次缺口</th><th>v{ver7} 累計外部資金</th><th>v{ver7} 首次缺口</th><th>v{ver7} 谷底</th><th>v{ver7}＋R4 累計外部資金</th><th>＋R4 首次缺口</th></tr></thead><tbody>{mc7_html}</tbody></table></div>
<h3>R5c 未歸屬餘額付款期：依 S-1/A 結構合成（判斷題「未歸屬餘額付款期」結案）</h3>
<div class="tw"><table><thead><tr><th>$B／比例</th>{"".join(f"<th>{y}</th>" for y in YEARS)}</tr></thead><tbody>{shr_html}</tbody></table></div>
<p class="note">做法：三段合計 {n("COST_S1fCommit3", None, "f1")} 的逐年付款＝表內義務逐年（2026Q4 {n("SRC_ANT_497", None, "f2")}、2027–2030 依原文表、之後 {n("SRC_ANT_502", None, "f1")} 依 2031–2036 平均〔Assumed〕）＋未起租 {n("SRC_ANT_504", None, "f1")} 與 9-30 後 {n("SRC_ANT_506", None, "f1")} 依「起租 2027、加權租期 {n("SRC_ANT_505", None, "f0")} 年」平均〔Assumed〕；除以三段合計得年度比例，套到未歸屬餘額 {n("COST_S1Unalloc", None, "f1")}。2025–2030 比例合計 {n(_r7("Cost", "R5c 2025–2030 比例合計"), None, "pct0")}（舊讀法 4／7＝57%）。新增 Checks：模型 2027–2030 合約實付（不含 AMD）對 S-1/A 表內＋租約推算 {n(_r7("Cost", "R5c 新 Checks：S-1/A 推算"), None, "f1")} 為 {n("COST_S1aPay2730Gap", None, "pct0")}——模型逐家合約的付款比 S-1/A 付款表早（逐家合約以合約期平均與爬坡攤提），只對照。並列讀法「殘差法」（未歸屬餘額逐年＝S-1/A 三段逐年 − 模型逐家合約逐年付款，獨立核對建議）把 2027–2028 未歸屬付款由 {n("COST_Pay_S1Unalloc", 2027, "f1")}／{n("COST_Pay_S1Unalloc", 2028, "f1")} 降到 {n("COST_S1aUnallocResid", 2027, "f1")}／{n("COST_S1aUnallocResid", 2028, "f1")}（命題輸出見下表）；基準不採用：上限型合約（SpaceXAI、Nscale）的模型付款可能超過最低承諾，殘差法會低估未歸屬付款，故基準取保守。</p>
<div class="tw"><table><thead><tr><th>讀法</th><th>未歸屬餘額付款 2027</th><th>2028</th><th>累計外部資金</th><th>首次缺口年</th><th>現金谷底</th><th>2030 差額</th><th>命題 1 持續轉正年</th></tr></thead><tbody>{alt7_html}</tbody></table></div>
<h3>既有預設的替代讀法，在 v{ver7} 基準下（每項只改一處；v0.1.7.1 當時的同一張表見 ①之六）</h3>
<div class="tw"><table><thead><tr><th>讀法</th><th>累計外部資金</th><th>首次缺口年</th><th>現金谷底</th><th>2026 差額</th><th>2030 差額</th></tr></thead><tbody>{alt8_html}</tbody></table></div>
<h3>R5a 2026 Q3 實績（S-1/A 季度損益；$B）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>S-1/A 實際</th><th>v0.1.7.1 用的報導或隱含值</th></tr></thead><tbody>{q3_html}</tbody></table></div>
<p class="note">W4 下半年非算力（含通路平台費、不含股權報酬與捐贈）改用 Q3 實際，Q4＝Q3 慣例不變；股權報酬代理下半年改（1.070＋1.960）× 2。非現金捐贈 1.960 為一次性（10 月以 Class N 股份結清），依 Q4＝Q3 慣例 Q4 也計一次，屬保守（判斷題 s1a_chr_q4_switch）；只影響 2026 含股權報酬的全成本、不影響現金，也不進每人股權報酬基準（獨立核對後修正：先前誤把下半年捐贈 3.92 帶進每人基準、沿用到 2027–2030）。只計 Q3 一次 → 2026 差額 {n("COST_PropGap_VR", 2026, "sf2", "chr_once")}、2026 覆蓋率 {n("COST_Coverage", 2026, "f3", "chr_once")}，2027–2030 不變。Q3 算力 9.863 與模型 Q3 隱含（報導營收 ÷ 1.70）差 {n("COST_S1aQ3ComputeGap", None, "sf3")}，只對照。調整後營業利益率原文 15%（取代報導 11%）。</p>
<h3>R5b 承諾總額（S-1/A 第 173–174 頁；$B）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>金額</th><th>基準處理</th></tr></thead><tbody>{cm_html}</tbody></table></div>
<p class="note">W1(b) 三段：S-1 的 6-30 口徑 {n(_r7("Cost", "R5b 對照：v0.1.7.1 三段"), None, "f1")} → S-1/A 的 9-30 口徑 {n("COST_S1fCommit3", None, "f1")}；逐家 Google 114.5 → {n("SRC_ANT_513", None, "f1")}（5–10 月修訂）、Microsoft 期間到 2033-06（金額不變；年度模型不受影響）。未歸屬餘額 {n("COST_S1Unalloc", None, "f1", "v0171")} → {n("COST_S1Unalloc", None, "f1")}。不在總額內、維持不進基準：無固定金額的其他算力安排 150–175（含 Google 額外 TPU 125–150 與 AMD >20）、可再發行可轉債最多 {n("SRC_ANT_512", None, "f0")}（Broadcom 42＋Google 20＋Amazon 15；N3）。</p>
<h3>R4 SpaceXAI：S-1 路徑（並列；基準維持月費 1.25 × 36＝45）</h3>
<div class="tw"><table><thead><tr><th>$B</th>{"".join(f"<th>{y}</th>" for y in YEARS)}</tr></thead><tbody>{spx_html}</tbody></table></div>
<p class="note">S-1 路徑：2026-05～11 交付期間線性爬升（第 k 月＝3.1 × k ÷ 8〔Assumed〕）、2026-12 起每月 {n("SRC_ANT_515", None, "f1")}、累計以上限 {n("SRC_ANT_453", None, "f1")} 截頂（2029 年部分月份）。SpaceX 合約總額 45 → 120.5，未歸屬餘額相應減少 {n("COST_S1Unalloc", None, "f1")} → {n("COST_S1Unalloc", None, "f1", "r4_on")}（總承諾不變，只是付款提前到 2026–2029）。2027-01／03 後可 90 天通知取消，未模型化。</p>
<h3>R5e 2026-09-30 現金對照（只對照，不校準）；R5f Q3 無新股權</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>v{ver7}</th></tr></thead><tbody>{c930_html}</tbody></table></div>
<p class="note">S-1/A 前九月特別股發行收入 {n("SRC_ANT_521", None, "f3")} 與上半年相同（Q3 無新股權；評估至 11-02 亦無），模型已到位融資 {n("FND_Committed", 2026, "f2")} 不變，差 {n("FND_CommittedS1aGap", None, "sf3")}。前九月營業現金流 {n("SRC_ANT_518", None, "f3")}（Q3 約 {n(_r7("Funding", "對照：S-1/A Q3 營業現金流"), 2026, "f2")}）、資本支出 {n("SRC_ANT_519", None, "f3")}、租賃資產預付 {n("SRC_ANT_520", None, "f3")}（新科目）。9-30 差額 {n("FND_Cash0930Gap", None, "sf2")} 主要是受限現金 2.6 與自由現金流時點；其餘 {n("FND_Cash0930Resid", None, "sf2")}。Google 額外 TPU 125–150、AMD >20 仍不進基準；TK 快照維持 v21.7（Tokenomics v21.9 尚未合併）。</p>
{V.src("COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExtTrough", "COST_S1fCommit3", "COST_S1Unalloc", "COST_Pay_S1Unalloc", "COST_S1aUnallocShare", "COST_S1aQ3NonComp", "COST_S1aSpaceXPath", "COST_S1aPay2730Gap", "FND_Cash0930Gap", "FND_CommittedS1aGap")}
</section>""")
    n = n_cur7
    # ── ①之六 v0.1.7.1 更新了什麼（v0.1.8 起以 v0.1.7.1 當時的輸出呈現：基準＝v0171〔R5a–R5c 皆關〕，替代讀法與最低現金＝v0171_ 版本；內容與前版相同） ──
    ver6 = "0.1.7.1"
    V.default_sc = "v0171"
    _m6 = {"base": "v0171", **{k: "v0171_" + k for k in ['s1u_fast', 's1u_slow', 'spx_cap', 'gaddl_on', 'bcm_2026', 'amd_s1', 'only_w2_off', 'only_w4_off', 'w4_neither', 'w4_q3only', 'w4_sbconly', 'commit_off', 'commit_off_w1only', 'mc_lo', 'mc_hi']}}
    n_cur6 = n
    n = lambda ref, year=None, f="f1", sc=None, cls="": n_cur6(ref, year, f, _m6.get(sc, sc), cls)  # noqa: E731
    def _r6(sheet, prefix):
        ws = V.loc.wb[sheet]
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and b.startswith(prefix):
                return f"{sheet}!{ws.cell(r, 1).value}"
        raise KeyError(f"{sheet} 找不到列：{prefix}")
    v6_cols = [("v0.1.7", "v017"), ("V2 前（V1、V3、V4 已套用）", "v0171_v2only_off"), ("＋各輪改原文 30.2／65.2", "v0171_round_only"),
               (f"＋Google Series G-2 計入＝v{ver6}", "base"), ("說明：只計入 Google G-2（各輪仍 30／65）", "v0171_googl_only")]
    def v6row(lab, ref, yr, f, em=False, cols=v6_cols):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cols) + "</tr>")
    v6_rows = [v6row("2026 已到位融資 $B", "FND_Committed", 2026, "f2", True)]
    v6_rows += [v6row(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2") for y in (2026, 2027, 2028, 2029, 2030)]
    v6_rows += [v6row("覆蓋率 2030（倍）", "COST_Coverage", 2030, "f3"), v6row("命題 1 持續轉正年", "COST_GapPosSustainYear", None, "yt"),
                v6row("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", 2030, "f2", True), v6row("首次缺口年", "FND_FirstGapYear", None, "yt"),
                v6row("只靠已到位融資的現金谷底 $B（2029）", "FND_CashNoExtTrough", None, "f2", True), v6row("2029 最低現金 $B", "FND_MinCash", 2029, "f2"),
                v6row("營運自由現金流持續轉正年", "FND_FCFPosSustainYear", None, "yt")]
    mc_defs = [("3 個月（區間下緣）", "v017_mc_lo", "mc_lo"), (f"6 個月（基準）", "v017", "base"), ("12 個月（區間上緣）", "v017_mc_hi", "mc_hi")]
    mc_html = "".join(
        f'<tr class="{"em" if b == "base" else ""}"><td>{lab}</td><td>{n("FND_MinCash", 2029, "f2", b)}</td>'
        f'<td>{n("FND_ExtNeedCum", 2030, "f2", a)}</td><td>{n("FND_FirstGapYear", None, "yt", a)}</td><td>{n("FND_CashNoExtTrough", None, "f2", a)}</td>'
        f'<td>{n("FND_ExtNeedCum", 2030, "f2", b)}</td><td>{n("FND_FirstGapYear", None, "yt", b)}</td><td>{n("FND_CashNoExtTrough", None, "f2", b)}</td></tr>'
        for lab, a, b in mc_defs)
    c630 = [("S-1 6-30 現金、約當現金與短期投資", "SRC_ANT_444", None, "f2"), ("模型近似（2025 年底＋已到位＋2026 自由現金流 × 6/12）", _r6("Funding", "模型 2026-06-30 近似"), 2026, "f2"),
            ("差（S-1 − 模型）", "FND_Cash0630Gap", None, "sf2"), ("①股權收入：S-1 特別股收入 − 模型已到位", _r6("Funding", "差額拆解 ①"), 2026, "sf3"),
            ("②上半年自由現金流時點", _r6("Funding", "差額拆解 ②"), 2026, "sf2"), ("③受限現金（Broadcom 租約保證金）", _r6("Funding", "差額拆解 ③"), 2026, "sf2"),
            ("④其餘", "FND_Cash0630Resid", None, "sf3")]
    c630_html = "".join(f"<tr><td>{lab}</td><td>{n(ref, yr, f, 'v017')}</td><td>{n(ref, yr, f)}</td></tr>" for lab, ref, yr, f in c630)
    alt6 = [(f"基準（v{ver6}）", "base"), ("未歸屬餘額付款期 快：2027–2031", "s1u_fast"), ("未歸屬餘額付款期 慢：2028–2036", "s1u_slow"),
            ("SpaceXAI 以正式 S-1 上限 120.5 計", "spx_cap"), ("Google 額外 TPU 125–150 計入（中點）", "gaddl_on"), ("Broadcom 起始年改 S-1 期間起點 2026", "bcm_2026"),
            ("AMD 改讀正式 S-1 下限 20", "amd_s1"), ("只關閉 W2（自建回到 Fluidstack $50B 三年均攤）", "only_w2_off"), ("只關閉 W4（非算力成本回到 v0.1.6 算法）", "only_w4_off"),
            ("W4 兩者皆不用（下半年 Q2 持平、股權報酬模型值）", "w4_neither"), ("W4 只動 Q3 口徑", "w4_q3only"), ("W4 只動股權報酬", "w4_sbconly"),
            ("承諾不進基準（W1、W2 皆關）", "commit_off"), ("說明列：只關 W1", "commit_off_w1only")]
    alt6_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("FND_ExtNeedCum", 2030, "f2", sc)}</td><td>{n("FND_FirstGapYear", None, "yt", sc)}</td>'
                        f'<td>{n("FND_CashNoExtTrough", None, "f2", sc)}</td><td>{n("COST_PropGap_VR", 2030, "sf2", sc)}</td></tr>' for lab, sc in alt6)
    S.append(f"""<section id="s1f"><h2><span class="no">①之六</span>v{ver6} 更新了什麼：正式 S-1 原文落地、已到位融資更正（Google Series G-2 約 10 另計）</h2>
<p class="lead"><b>命題 2 由「需要少量」翻回「不需要」（條件式）。</b>2026–2030 累計外部資金需求 {n("FND_ExtNeedCum", 2030, "f2", "v017")}（首次 {n("FND_FirstGapYear", None, "yt", "v017")} 年）→ {n("FND_ExtNeedCum", 2030, "f2")}；只靠已到位融資的現金谷底 {n("FND_CashNoExtTrough", None, "f2", "v017")} → {n("FND_CashNoExtTrough", None, "f2")}（2029 年；當年最低現金 {n("FND_MinCash", 2029, "f2")}，餘裕 {n("FND_Headroom", 2029, "f2")}）。原因只有 V2 一項（Google G-2 +10.0，另各輪改原文值 +0.4）：正式 S-1 合併現金流量表的 2026 上半年特別股發行收入為 {n("SRC_ANT_472", None, "f3")}，v0.1.7 的已到位融資只算 {n("FND_Committed", 2026, "f1", "v017")}——Google 2026 Q2 認購的 Series G-2（{n("SRC_ANT_475", None, "f1")} 百萬股 × 每股 {n("SRC_ANT_477", None, "f4")}≈{n("Funding!F17", 2026, "f2")}）與 2026-05 Google 投入 Series H 的 {n("SRC_ANT_480", None, "f1")} 是兩筆（Note 4、9、10），v0.1.7 誤當同一筆；Series G／H 另依原文改 {n("SRC_ANT_473", None, "f1")}／{n("SRC_ANT_478", None, "f1")}（＋0.4）。v{ver6} 已到位融資 {n("FND_Committed", 2026, "f2")}，與 S-1 差 {n("FND_CommittedS1Gap", None, "sf3")}。<b>命題 1 不變</b>（差額、覆蓋率、持續轉正年 {n("COST_GapPosSustainYear", None, "yt")} 全同）。這個翻轉幅度小、取決於最低現金假設：最低現金改 12 個月仍需 {n("FND_ExtNeedCum", 2030, "f2", "mc_hi")}（{n("FND_FirstGapYear", None, "yt", "mc_hi")} 年）。</p>
<h3>V2 瀑布：只動已到位融資（V1、V3、V4、V5 不改基準輸出）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in v6_cols)}</tr></thead><tbody>{"".join(v6_rows)}</tbody></table></div>
<p class="note">「V2 前」＝本版活頁簿只把 V2 的兩個開關（googl_upfront_switch、s1f_round_switch）設 0，數字與 v0.1.7 逐格相同：V1（46 列改 Verified、SRC_ANT_427–431 改原文精度）、V3（條件式額度）、V4（信用額度條款）、V5（文字）都不改基準輸出。逐格比對方法與格數見報告 ④。</p>
<h3>最低現金月數 3／6／12（D17：基準 6 個月，Assumed 3–12；V4）</h3>
<div class="tw"><table><thead><tr><th>最低現金</th><th>2029 最低現金</th><th>v0.1.7 累計外部資金</th><th>v0.1.7 首次缺口</th><th>v0.1.7 谷底</th><th>v{ver6} 累計外部資金</th><th>v{ver6} 首次缺口</th><th>v{ver6} 谷底</th></tr></thead><tbody>{mc_html}</tbody></table></div>
<p class="note">S-1 揭露的契約下限：A&R 循環信用額度（2026-09 擴為 {n("SRC_ANT_485", None, "f0")}，2031-09 到期）附最低流動性條款 {n("SRC_ANT_486", None, "f0")}（自 2026-12-31 季末起）。模型最低現金（次年非算力成本 × 6 個月）2026–2030 為 {n("FND_MinCash", 2026, "f1")}–{n("FND_MinCash", 2030, "f1")}，比契約下限高至少 {n("Checks!" + chk_of("h4_liq"), None, "f1")}；循環額度照 D16 只列或有，不作資金來源。</p>
<h3>v0.1.7 已套用預設的替代讀法，在 v{ver6} 基準下（每項只改一處）</h3>
<div class="tw"><table><thead><tr><th>讀法</th><th>累計外部資金</th><th>首次缺口年</th><th>現金谷底</th><th>2030 差額</th></tr></thead><tbody>{alt6_html}</tbody></table></div>
<p class="note">已到位融資多 10.4 後，「W4 只動 Q3 口徑」不再產生外部資金需求（v0.1.7 時為 6.29）；未歸屬餘額快付、SpaceXAI 上限、Google 額外 TPU、Broadcom 2026 起付、自建回到 Fluidstack 舊法等讀法仍會產生，只是各少約 10.4（v0.1.7 當時的同一張表見 ①之五）。</p>
<h3>V3 條件式額度（只動情境，基準不變）</h3>
<div class="tw"><table><thead><tr><th>輸出</th><th>V3 前（Google 30）</th><th>v{ver6}（Google 20）</th></tr></thead><tbody>
<tr><td>Google 條件式額度（2027 到位）</td><td>{n("Funding!F23", 2027, "f0", "v0171_v3_off")}</td><td>{n("Funding!F23", 2027, "f0")}</td></tr>
<tr><td>S2 ＋條件式：累計外部資金需求</td><td>{n("FND_ExtNeedCumCond", 2030, "f1", "v0171_v3_off")}</td><td>{n("FND_ExtNeedCumCond", 2030, "f1")}</td></tr>
<tr><td>S3 ＋IPO＋條件式：2030 年底現金</td><td>{n("FND_CashEndAll", 2030, "f1", "v0171_v3_off")}</td><td>{n("FND_CashEndAll", 2030, "f1")}</td></tr>
</tbody></table></div>
<p class="note">正式 S-1：Google 2026 NPA 於 2026-06-30 剩餘可發行 {n("SRC_ANT_482", None, "f0")}；另有 IPO 同時以發行價認購最多 {n("SRC_ANT_483_Hi", None, "f0")} 的選擇權，若行使，NPA 額度等額減少，兩者不相加。Amazon 2026 NPA 剩餘 {n("SRC_ANT_484", None, "f0")}（每達算力里程碑 5.0）。條件式一律不計入基準。</p>
<h3>2026-06-30 現金對照（只對照，不校準）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>v0.1.7</th><th>v{ver6}</th></tr></thead><tbody>{c630_html}</tbody></table></div>
<p class="note">V2 後「S-1 − 模型近似」為 {n("FND_Cash0630Gap", None, "sf2")}（模型較多），差額幾乎全部來自 Broadcom 租約保證金 {n("SRC_ANT_487", None, "f1")}（2026-04 存入的受限現金，不在 S-1「現金、約當現金與短期投資」內，之後抵付租金；預設不另計）與上半年自由現金流的時點；股權收入已與 S-1 一致。</p>
<p class="note"><b>V1 來源等級</b>：原為二手轉載的 46 列（季度損益、營收成本、承諾三段、逐家承諾、營業現金流與資本支出、集中度、non-GAAP 調節）改 Verified，網址改 SEC 原文、備註加原文位置；SRC_ANT_421 為轉載者捨入（原文表內 11,555、正文 11.6），只對照、標記不改；SRC_ANT_427–431 改原文精度（只對照列）。<b>V5 文字更正</b>：v0.1.7 報告與交接所寫「上半年股權報酬未單獨揭露」「約 0.111 未說明」改為：調整後 − GAAP 營業損益 2.311＝股權報酬 2.185（Q1 924＋Q2 1,261）＋非現金捐贈配對 0.126（2026Q1），兩者皆非現金，股權報酬代理口徑成立、數值不變。</p>
{V.src("FND_Committed", "FND_CommittedS1Gap", "FND_ExtNeedCum", "FND_CashNoExtTrough", "FND_MinCash", "FND_Cash0630Gap", "FND_Cash0630Resid", "FND_MinCashOverCovenant", "FND_ExtNeedCumCond")}
</section>""")
    n = n_cur6

    # ── ①之五 v0.1.7 更新了什麼（v0.1.7.1 起以 v0.1.7 當時的輸出呈現：基準＝v017〔V2、V3 皆關〕，替代讀法＝v017_ 版本；內容與前版相同） ──
    V.default_sc = "v017"
    _m5 = {"base": "v017", **{k: "v017_" + k for k in ['commit_off', 'commit_off_w1only', 's1u_fast', 's1u_slow', 'spx_cap', 'gaddl_on', 'bcm_2026', 'amd_s1', 'only_w2_off', 'only_w4_off', 'w4_neither', 'w4_q3only', 'w4_sbconly']}}
    n_cur = n
    n = lambda ref, year=None, f="f1", sc=None, cls="": n_cur(ref, year, f, _m5.get(sc, sc), cls)  # noqa: E731
    ver5 = "0.1.7"
    def _r(sheet, prefix):
        ws = V.loc.wb[sheet]
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and b.startswith(prefix):
                return f"{sheet}!{ws.cell(r, 1).value}"
        raise KeyError(f"{sheet} 找不到列：{prefix}")
    wf_cols = [("v0.1.6", "v016"), ("＋W1 承諾進基準", "wf_w1"), ("＋W2 自建改 GW 法", "wf_w2"), ("＋W3 2025 校準", "wf_w3"),
               ("＋W4 非算力校準", "wf_w4"), (f"＋W5 快照同步＝v{ver5}", "base")]
    def wfrow(lab, ref, yr, f, em=False, cols=wf_cols):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cols) + "</tr>")
    wf_rows = [wfrow(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", y == 2030) for y in (2026, 2027, 2028, 2029, 2030)]
    wf_rows += [wfrow("覆蓋率 2030（倍）", "COST_Coverage", Y, "f3"), wfrow("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1", True),
                wfrow("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2"), wfrow("算力成本 2028 $B", "COST_Compute", 2028, "f1"),
                wfrow("供給 VR 等值 GW 2028", "CMP_Supply_VReq", 2028, "f2"), wfrow("非算力成本（不含股權報酬）2026 $B", "COST_NonCompExSBC", 2026, "f2"),
                wfrow("η（2025 校準）", "CMP_Eta2025", None, "f3")]
    two_cols = [(f"承諾進基準（v{ver5} 基準）", "base"), ("不進基準（v0.1.6 合約口徑：W1、W2 皆關）", "commit_off"), ("說明列：只關 W1（缺陷：Fluidstack 不計任何成本）", "commit_off_w1only")]
    two_rows = [wfrow(f"每 VR 等值 GW 差額 {y}", "COST_PropGap_VR", y, "sf2", y == 2030, two_cols) for y in (2026, 2027, 2028, 2029, 2030)]
    two_rows += [wfrow(f"覆蓋率 {y}", "COST_Coverage", y, "f3", False, two_cols) for y in (2026, 2027, 2028, 2029, 2030)]
    two_rows += [wfrow("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1", True, two_cols),
                 wfrow("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2", False, two_cols),
                 wfrow("首次缺口年", "FND_FirstGapYear", None, "yt", False, two_cols)]
    pp_defs = [("Google Cloud（至少）", "Cost!K121", "SRC_ANT_450"), ("Google／Broadcom TPU 租賃", "Cost!K122", "SRC_ANT_454"), ("Amazon（至少）", "Cost!K123", "SRC_ANT_451"),
               ("Microsoft Azure（至少）", "Cost!K124", "SRC_ANT_452"), ("SpaceXAI（最高；沿用月費）", "Cost!K126", "SRC_ANT_453"), ("Nscale（最高；上限全數計入）", "Cost!K127", "SRC_ANT_455"),
               ("Lambda（約）", "Cost!K128", "SRC_ANT_456"), ("Volta Infra", "Cost!K129", None), ("Akamai", "Cost!K130", None), ("AMD（不在 S-1 承諾表內；不扣；維持模型 GW 型）", "Cost!K125", "SRC_ANT_457")]
    pp_html = "".join(f"<tr><td>{lab}</td><td>{n(t, 2025, 'f1', 'v016')}</td><td>{n(t, 2025, 'f1')}</td><td>{n(src, None, 'f1') if src else '—'}</td></tr>" for lab, t, src in pp_defs)
    pp_html += (f'<tr class="em"><td>模型逐家合計（不含 AMD）</td><td>—</td><td>{n("COST_S1fCommitModel", None, "f1")}</td><td>三段合計 {n("COST_S1fCommit3", None, "f1")}</td></tr>'
                f'<tr class="em"><td>S-1 承諾未歸屬餘額（三段合計 − 模型逐家）</td><td>0</td><td>{n("COST_S1Unalloc", None, "f1")}</td><td>—</td></tr>')
    alt_cols = [("2027 差額", 2027, "COST_PropGap_VR", "sf2"), ("2030 差額", Y, "COST_PropGap_VR", "sf2"), ("2030 覆蓋率", Y, "COST_Coverage", "f3"),
                ("累計外部資金", Y, "FND_ExtNeedCum", "f1"), ("現金谷底", None, "FND_CashNoExtTrough", "f2")]
    alt_defs = [(f"基準（v{ver5}）", "base"), ("未歸屬餘額付款期 快：2027–2031", "s1u_fast"), ("未歸屬餘額付款期 慢：2028–2036", "s1u_slow"),
                ("SpaceXAI 以正式 S-1 上限 120.5 計", "spx_cap"), ("Google 額外 TPU 125–150 計入（中點）", "gaddl_on"), ("Broadcom 起始年改 S-1 期間起點 2026", "bcm_2026"),
                ("AMD 改讀正式 S-1 下限 20", "amd_s1"), ("只關閉 W2（自建回到 Fluidstack $50B 三年均攤）", "only_w2_off"), ("只關閉 W4（非算力成本回到 v0.1.6 算法）", "only_w4_off")]
    w4_defs = [("W4 兩者皆不用（v0.1.7 首版：下半年 Q2 持平、股權報酬模型值）", "w4_neither"), ("只動 Q3 口徑（下半年與營收同期）", "w4_q3only"),
               ("只動股權報酬（S-1 調整數代理）", "w4_sbconly"), (f"兩者（v{ver5} 基準）", "base")]
    w4_cols = [("2026 非算力（不含股權報酬）", 2026, "COST_NonCompExSBC", "f2"), ("2026 股權報酬", 2026, "COST_SBC", "f2"), ("2030 差額", Y, "COST_PropGap_VR", "sf2"),
               ("2030 覆蓋率", Y, "COST_Coverage", "f3"), ("累計外部資金", Y, "FND_ExtNeedCum", "f1"), ("首次缺口年", None, "FND_FirstGapYear", "yt"), ("現金谷底", None, "FND_CashNoExtTrough", "f2")]
    w4_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _h, yr, ref, f in w4_cols) + "</tr>" for lab, sc in w4_defs)
    alt_html = "".join(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _h, yr, ref, f in alt_cols) + "</tr>"
                       for lab, sc in alt_defs)
    nnv = [("TPU v7 持有比", "TK_NNV_TPUv7_Hold"), ("Trainium3 持有比", "TK_NNV_Trn3_Hold"), ("MI455X 持有比", "TK_NNV_MI455X_Hold"), ("MI455X 產出比", "TK_NNV_MI455X_Out")]
    nnv_html = "".join(f"<tr><td>{lab}</td><td>{n(nm_, None, 'f2', 'v016')}</td><td>{n(nm_, None, 'f2')}</td></tr>" for lab, nm_ in nnv)
    S.append(f"""<section id="s1e"><h2><span class="no">①之五</span>v{ver5} 更新了什麼：正式 S-1 全面套用、自建資本支出改 GW 法、Tokenomics v21.7 與 OpenAI v0.6.4 快照</h2>
<p class="lead"><b>結論部分翻轉。</b>命題 1：每 VR 等值 GW 差額由 v0.1.6 的「2026–2028 為負、2029 起轉正」，變為「2026 為正、2027–2029 為負、2030 才轉正」（2029 年 {n("COST_PropGap_VR", 2029, "sf2", "v016")} → {n("COST_PropGap_VR", 2029, "sf2")}；2030 年 {n("COST_PropGap_VR", Y, "sf2", "v016")} → {n("COST_PropGap_VR", Y, "sf2")}；2030 覆蓋率 {n("COST_Coverage", Y, "f3", "v016")} → {n("COST_Coverage", Y, "f3")}）。命題 2 翻轉為「需要少量」：累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "v016")} → {n("FND_ExtNeedCum", Y, "f2")}（首次 {n("FND_FirstGapYear", None, "yt")} 年），只靠已到位融資的現金谷底由 {n("FND_CashNoExtTrough", None, "f1", "v016")} 降到 {n("FND_CashNoExtTrough", None, "f1")}；多個單一驅動落在下緣缺口就大幅擴大（⑥）。最大的兩個變動方向相反：W1 把正式 S-1 承諾（含 6-30 後新簽 245）放進基準，2027 年起每年多付約 {n("COST_Pay_S1Unalloc", 2028, "f1")}；W2 依 Andy 10-31 指示改用自建 GW × TK 每 GW 建置成本，Wattonomics 沒有 Anthropic 自有 GW 證據，自建資本支出由每年 {n(_r("Cost", "W2 對照：舊法"), 2026, "f1")} 變為 0。</p>
<h3>瀑布：從 v0.1.6 逐項加入（W6 只核對，不改數值）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in wf_cols)}</tr></thead><tbody>{"".join(wf_rows)}</tbody></table></div>
<h3>兩讀並列：正式 S-1 承諾進基準 vs 不進基準（工作單 W1(d)；不進基準＝W1、W2 皆關，因 W2 只在承諾進基準時自洽——Fluidstack 改由承諾中的租回承擔；第三欄只關 W1 會讓 Fluidstack 既無資本支出也無租金，只供說明）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in two_cols)}</tr></thead><tbody>{"".join(two_rows)}</tbody></table></div>
<h3>W1 承諾進基準：逐家改讀正式 S-1，未歸屬餘額以通用合約價換 GW（$B）</h3>
<div class="tw"><table><thead><tr><th>合約</th><th>v0.1.6 模型總額</th><th>v{ver5} 模型總額</th><th>正式 S-1</th></tr></thead><tbody>{pp_html}</tbody></table></div>
<p class="note">S-1 三段（6-30 尚未起租租賃 {n("SRC_ANT_445", None, "f1")}＋6-30 固定或最低採購承諾 {n("SRC_ANT_446", None, "f1")}＋6-30 後新簽 {n("SRC_ANT_447", None, "f0")}；與 Note 16 的 {n("SRC_ANT_448", None, "f0")} 重疊，不相加）扣除模型逐家合約總額後的未歸屬餘額 {n("COST_S1Unalloc", None, "f1")}，以通用合約價 {n("INP_114", None, "f0")}（VR200 族）換 GW、2027–2033 平均攤提：每年付款 {n("COST_Pay_S1Unalloc", 2028, "f2")}、供給 {n("CMP_Supply_S1Unalloc", 2028, "f2")} GW，2030 年後仍應付 {n(_r("Cost", "W1(b) 未歸屬餘額中 2030"), None, "f1")}。付款與 GW 的換算沿用既有方法（金額型 GW＝付款 ÷ 合約價；Broadcom 雙揭露型供給維持 8-K 的 3.5 GW、付款改 161.2）；AMD「不在承諾表內」，維持「最多 2 GW × 計入比例」。SpaceXAI 與 Nscale 屬上限型：SpaceX 沿用月費（上限 120.5 只作區間，差額落在未歸屬餘額），Nscale 沿用 U2 的上限全數計入。Google 額外 TPU 125–150 不在承諾表，只作情境（下表）。</p>
<h3>W2 自建資本支出：自建 GW 增量 × TK 每 GW 建置成本（廠房依 18 個月提前入帳）</h3>
<p class="note">自建 GW 以 Wattonomics v0.1.7 為權威（各年 0：沒有 Anthropic 自有站點 GW 證據；未找到不等於不存在，列為資料缺口）。Fluidstack 站點視為 Fluidstack 擁有、經承諾租回（已在 W1 的 S-1 承諾內），不再另計。模型 2026 自建資本支出 {n("COST_OwnedCapex", 2026, "f2")}，S-1 上半年資本支出年化 {n(_r("Cost", "W2 對照：S-1 2026"), 2026, "f2")}（含辦公室等非算力項目，只作對照）；舊法（Fluidstack $50B 三年均攤）2026–2028 每年 {n(_r("Cost", "W2 對照：舊法"), 2026, "f1")}。</p>
<h3>W3 2025 校準改正式值、W4 2026 非算力成本以正式季度營業損益校準</h3>
<p class="note">W3：2025 營收 {n("SRC_ANT_001", None, "f2")} → {n("SRC_ANT_412", None, "f3")}、算力與基礎設施費用 {n("SRC_ANT_252", None, "f2")} → {n("SRC_ANT_424", None, "f3")}，η 由 {n("CMP_Eta2025", None, "f4", "wf_w2")} 重校為 {n("CMP_Eta2025", None, "f4", "wf_w3")}，影響極小。W4（獨立核對後修正）：S-1 2026 上半年非算力營業費用＝營收 − 算力與基礎設施 − 營業損益＝{n(_r("Cost", "W4 2026 上半年非算力"), 2026, "f3")}（含平台抽成與股權報酬）。股權報酬改用 S-1「調整後 − GAAP 營業損益」代理：上半年 {n(_r("Cost", "股權報酬代理：2026 上半年"), 2026, "f3")}、Q2 {n(_r("Cost", "股權報酬代理：2026 Q2"), 2026, "f3")}，全年照營收同法（Q3＝Q2、Q4＝Q3）＝{n("COST_S1fSBCProxy2026", None, "f2")}（舊模型值 {n(_r("Cost", "對照：v0.1.6 算法的 2026 股權報酬"), 2026, "f2")}；S-1 人事附註上半年股權報酬 {n("SRC_ANT_467", None, "f1")}，只對照）。下半年改與營收同期：Q3＝Q3 報導隱含非算力 {n(_r("Cost", "報導隱含：Q3 非算力"), 2026, "f2")}（含抽成、不含股權報酬）、Q4＝Q3，下半年 {n(_r("Cost", "W4 2026 下半年非算力"), 2026, "f2")}。全年（不含股權報酬）{n(_r("Cost", "W4 2026 全年非算力"), 2026, "f2")} − 平台抽成 {n(_r("Cost", "W4 扣：模型 2026 雲端"), 2026, "f2")}＝<b>{n("COST_S1fNonComp2026", None, "f2")}</b>，比 v0.1.6 算法 {n(_r("Cost", "W4 對照：v0.1.6 算法"), 2026, "f2")} 高 {n("COST_S1fNonCompDiff", None, "pct0")}（超過 30% 門檻：照推得值入基準，並列舊值）。交叉核對：同法算出的 S-1 2025 非算力 − 模型 2025 值＝{n(_r("Cost", "W4 交叉核對：S-1 2025 非算力營業費用（營收"), 2025, "f3")}。兩項修正方向相反，各自與合併的影響如下表（只動 Q3 口徑時第 2 個結論翻轉）。</p>
<div class="tw"><table><thead><tr><th>W4 讀法</th>{"".join(f"<th>{h}</th>" for h, *_r2 in w4_cols)}</tr></thead><tbody>{w4_html}</tbody></table></div>
<h3>W5 快照同步、W6 Amazon 實際投入</h3>
<div class="tw"><table><thead><tr><th>Tokenomics NonNV 比例（進計算者）</th><th>v21.2</th><th>v21.7</th></tr></thead><tbody>{nnv_html}</tbody></table></div>
<p class="note">W5：TK_Link 改 Tokenomics {html.escape(str(tk_ver))}（29 個值變動：17 個具名範圍＋NonNV 12 格；另新增 IF_CapexIT、IF_CapexFacility 供 W2）。進計算的主要是 TPU v7、Trainium3 持有比下降：既有雲端加權合約價變低，同額付款換得較多 GW，η（2025）{n("CMP_Eta2025", None, "f3", "wf_w4")} → {n("CMP_Eta2025", None, "f3")}（低於 Checks 的合理下緣 0.5，C54 警示），2025、2026 年底 GW 對報導偏高（C69、C70 警示）；命題輸出變動小（2030 差額 {n("COST_PropGap_VR", Y, "sf2", "wf_w4")} → {n("COST_PropGap_VR", Y, "sf2")}）。OAI_Link 改 OpenAI v0.6.4（2026 每 VR 等值 GW 營收 {n("OAI_COST_PropRev_VR", 2026, "f2")}；只並排）。W6：Amazon 2026Q3 10-Q 顯示 2026 年實際投入 {n(_r("Funding", "Amazon 10-Q"), 2026, "f1")}（Series G 5＋Series H 5）；模型已到位融資中 Amazon 部分同為 {n(_r("Funding", "模型：2026 已到位融資中的 Amazon"), 2026, "f1")}（Series G 特別股 5＋Series H $65B 內含 5），差 {n("FND_AmznCheck2026", 2026, "f1")}——模型從未以 18.0（含 2023–2025 可轉債 8）計入 2026 流入，18.0 只用在回流對照的累計欄，不改。</p>
<h3>已套用預設的替代讀法（每項只改一處）</h3>
<div class="tw"><table><thead><tr><th>讀法</th>{"".join(f"<th>{h}</th>" for h, *_r2 in alt_cols)}</tr></thead><tbody>{alt_html}</tbody></table></div>
{V.src("COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExtTrough", "COST_S1Unalloc", "COST_Pay_S1Unalloc", "CMP_Supply_S1Unalloc", "COST_S1fNonComp2026", "COST_OwnedCapex", "FND_AmznCheck2026")}
</section>""")
    n = n_cur
    V.default_sc = "v016"         # ①之二～①之四：以 v0.1.6 當時的輸出呈現（內容與前版相同）
    # ── ①之二 v0.1.1 更新了什麼 ──
    cols = [("v0.1（更新前）", "v010"), ("只加 U1（2026 營收）", "only_u1"), ("只加 U2（Nscale＋Lambda）", "only_u2"), ("v0.1.1（更新後）", "v011")]
    def upd(lab, ref, yr, f, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in cols) + "</tr>")
    upd_rows = [upd("2026 營收總額（報導口徑）$B", "REV_GrossCapped", 2026, "f1", True)]
    upd_rows += [upd(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", y == 2030) for y in (2026, 2027, 2028, 2029, 2030)]
    upd_rows += [upd("覆蓋率 2030", "COST_Coverage", Y, "pct0"), upd("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1", True),
                 upd("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f1")]
    tk_cols = [("v0.1.1（TK v5.26）", "v011"), ("只換 IF_TokGW 群", "only_tokgw"), ("只換 IF_HoldEcon", "only_hold"), (f"v0.1.2（TK v18.0；v19.0 值相同）", "v012")]
    def tkrow(lab, ref, yr, f, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in tk_cols) + "</tr>")
    tk_rows = [tkrow("η（2025 校準；各年同值）", "CMP_Eta", 2025, "f3", True), tkrow("供給 VR 等值 GW 2030", "CMP_Supply_VReq", Y, "f2")]
    tk_rows += [tkrow(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", y == 2030) for y in (2025, 2026, 2027, 2028, 2029, 2030)]
    tk_rows += [tkrow("覆蓋率 2030（倍）", "COST_Coverage", Y, "f3"), tkrow("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1", True),
                tkrow("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2"),
                tkrow("供應商持有成本 2030 $B（IF_HoldEcon × GW；只作參考）", "COST_SupplierHold", Y, "f1"), tkrow("雲端毛利 2030 $B（合約實付 − 持有成本）", "COST_CloudGM", Y, "f1")]
    pq_rows = [
        ("TK L1_RevGW_PQ_VR200：VR200 每 GW 營收 P × Q（第一年；m＝2026 Q2）", "Cost!K240", "f1"), ("　低（實際利用率 0.30）", "Cost!K241", "f1"),
        ("　高（實際利用率 0.65）", "Cost!K242", "f1"), ("　m 改用 FY2026", "Cost!K243", "f1"), ("　壽命期年均（第一年 × L1_LTotalDisc）", "Cost!K244", "f1"),
        ("本模型：2026 營收總額 ÷ 付費推論 VR 等值 GW", "Cost!K246", "f1"), ("本模型：2026 營收淨額 ÷ 供給 VR 等值 GW（命題 1 的分子口徑）", "Cost!K249", "f1"),
        ("TK P × Q ÷ 本模型（付費推論口徑）", "Cost!K250", "f2"), ("TK P × Q ÷ 本模型（供給口徑，含閒置與研發 GW；僅供參考）", "Cost!K251", "f2"),
        ("TK token 帳未閉合倍數（模型付費 token ÷ 需求面付費 token）", "Cost!K245", "f2"), ("本模型每 GW 付費 token ÷ TK Q", "Cost!K255", "f2")]
    pq_html = "".join(f"<tr><td>{lab}</td><td>{n(ref, 2026, f)}</td></tr>" for lab, ref, f in pq_rows)
    S.append(f"""<section id="s1b"><h2><span class="no">①之二</span>v0.1.2 更新了什麼：Tokenomics v18.0（v0.1.3 快照 v19.0 值相同；v0.1.4 改 v19.1，見 ①之三）</h2>
<p class="lead">Tokenomics 由 v5.26 升到 v18.0 後，命題結論不變：2026–2028 仍為負、2029 起轉正，累計外部資金需求仍為 {n("FND_ExtNeedCum", Y, "f0")}。
變動集中在 η：TK 每 GW 年產能（IF_TokGW_*）上升約 1.3–1.8 倍（v9.0 快取命中進產能、v10.2 每用戶速度 SLO 改量測值），η 由 {n("CMP_Eta", 2025, "f3", "v011")} 降到 {n("CMP_Eta", 2025, "f3")}，把整體上升吸收掉，所以供給 VR 等值 GW 2030 只由 {n("CMP_Supply_VReq", Y, "f2", "v011")} 變為 {n("CMP_Supply_VReq", Y, "f2", "v012")}（v0.1.3 再因 U2 新法改為 {n("CMP_Supply_VReq", Y, "f2", "v013")}），2026–2030 各年每 VR 等值 GW 差額變動不超過 0.1（2025 年分母只有約 0.06 VR 等值 GW，該年數字本來就極端）。
公式沒有改；只更正五處過時的說明文字，見報告「語意核對」。</p>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in tk_cols)}</tr></thead><tbody>
{"".join(tk_rows)}
</tbody></table></div>
<p class="note">「只換」欄＝其他 TK 名稱維持 v5.26 值，只把該群換成 v18.0；TK 另有 26 個值變動的名稱（IF_Alloc*、IF_FullCost* 等）本模型只快照、不進命題輸出。IF_CapexTotal 的 VR200 欄未變，自建 GW 不受影響。逐名稱新舊值與變動版本見報告。</p>
<h3>與 TK v18.0 新名稱對照：每 VR200 GW 營收 P × Q（只對照，不回饋計算）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>2026</th></tr></thead><tbody>{pq_html}</tbody></table></div>
<p class="note">TK 自己標明 token 帳未閉合：模型付費 token 約為需求面的 {n("Cost!K245", 2026, "f1")} 倍，P × Q 是「模型產能推算的 Q × 需求面 P」，若誤差全在產能側，實際營收約為該列除以此倍數。兩邊口徑不同（TK 為 VR200 每 GW 第一年、取實際利用率；本模型以 η 校準到 2025 推論支出），請當作量級對照，不要相減。</p>
{V.src("CMP_Eta", "CMP_Supply_VReq", "COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExtTrough", "COST_SupplierHold", "COST_CloudGM", "COST_TKPQ", "COST_TKPQRatio", "COST_TKTokLedger")}
<details><summary><b>v0.1.1 更新摘要（S-1 草稿與 Q3 簡報；U1–U8）</b>：2026 營收上修、Nscale＋Lambda 錨定 86，命題結論不變</summary>
<p class="lead">草擬 S-1（WSJ 2026-10-21 轉述）與 Q3 投資人簡報進來後，命題結論不變：2026–2028 仍為負、2029 起轉正，累計外部資金需求仍為 {n("FND_ExtNeedCum", Y, "f0")}。2026 營收由 {n("REV_GrossCapped", 2026, "f1", "v010")} 上修到 {n("REV_GrossCapped", 2026, "f1")}（U1），2026 年每 VR 等值 GW 差額由 {n("COST_PropGap_VR", 2026, "sf2", "v010")} 收斂到 {n("COST_PropGap_VR", 2026, "sf2", "v011")}；Nscale＋Lambda 改錨定『最高 $86B』（U2）使 2027 起算力付款與供給 GW 變大，2028 年差額由只加 U1 時的 {n("COST_PropGap_VR", 2028, "sf2", "only_u1")} 變為 {n("COST_PropGap_VR", 2028, "sf2", "v011")}（v0.1 為 {n("COST_PropGap_VR", 2028, "sf2", "v010")}）。其餘更新項（U3–U5、U7）只做對照，不改任何基準數字；談判中的 $150B（U6）只是情境，基準仍只計已簽合約。</p>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in cols)}</tr></thead><tbody>
{"".join(upd_rows)}
</tbody></table></div>
<p class="note">「更新前」＝把 Inputs 的 U1、U2 開關關掉、並把 TK 快照回到 v5.26，Excel 重算的結果（等於 v0.1 的輸出）；「更新後」同樣用 TK v5.26（即 v0.1.1 當時的數字）；兩欄「只加」用來拆解各項貢獻。U3–U8 沒有基準數字變動：它們是與報導的對照列（見下）。</p>
<h3>與報導對照（U3、U4、U5；只對照，不校準）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>模型</th><th>報導</th><th>差距</th></tr></thead><tbody>
<tr><td>U3 雲端付款義務（約）$B</td><td>{n("Cost!K190", None, "f0")}</td><td>{n("Cost!K191", None, "f0")}</td><td>{n("COST_CommitGapCloud", None, "pct1")}</td></tr>
<tr><td>U3 已鎖定承諾（>，含 Broadcom 租賃）$B</td><td>{n("Cost!K193", None, "f0")}</td><td>{n("Cost!K194", None, "f0")}</td><td>{n("COST_CommitGapLocked", None, "pct1")}</td></tr>
<tr><td>U3 含談判中（約）$B</td><td>{n("Cost!K196", None, "f0")}</td><td>{n("Cost!K197", None, "f0")}</td><td>{n("COST_CommitGapAll", None, "pct1")}</td></tr>
<tr><td>U4 2026 算力：模型合約實付 vs 報導外推全年 $B（上半年 {n("Cost!K204", 2026, "f0")}＋Q3＋Q4）</td><td>{n("Cost!K208", 2026, "f1")}</td><td>{n("Cost!K207", 2026, "f1")}</td><td>{n("COST_SpendGap2026Pay", None, "pct1")}</td></tr>
<tr><td>U4 同上，模型含自建資本支出（Fluidstack）</td><td>{n("Cost!K209", 2026, "f1")}</td><td>{n("Cost!K207", 2026, "f1")}</td><td class="neg">{n("COST_SpendGap2026Cc", None, "pct1")}</td></tr>
<tr><td>U5 2026 調整後營業利益率：模型 vs Q3 報導</td><td>{n("Cost!K216", 2026, "pct1")}（差，百分點）</td><td>{n("COST_AdjOIMRep2026Q3", None, "pct0")}</td><td>—</td></tr>
<tr><td>U5 隱含 2026 非算力成本低估 $B（Q3 年化法／營收比例法）</td><td>{n("COST_D22UnderQ3", None, "f1")}</td><td>{n("COST_D22UnderQ3Ratio", None, "f1")}</td><td>—</td></tr>
</tbody></table></div>
<p class="note">U4 含自建的差距超過 {n("INP_183", None, "pct0")} 門檻，Checks C116 為 WARN：Fluidstack 自建 16.7 是現金資本支出，是否與報導的『算力與基礎設施支出』同口徑，S-1 摘要沒有說明。U7：上半年約 72 的非現金費用不進現金流（見 Funding）。U8：2026-06-30 現金餘額在轉述中找不到，列為資料缺口。</p>
<h3>U6 情境：談判中的新協議約 $150B（不進基準）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>2028</th><th>2029</th><th>2030</th></tr></thead><tbody>
<tr><td>新協議付款 $B（從 {n("Cost!K225", None, "yr")} 起，{n("Cost!K226", None, "f0")} 年，爬坡）</td>{"".join(f"<td>{n('COST_NegPay', y, 'f1')}</td>" for y in (2028, 2029, 2030))}</tr>
<tr><td>新增供給 GW</td>{"".join(f"<td>{n('CMP_NegGW', y, 'f2')}</td>" for y in (2028, 2029, 2030))}</tr>
<tr class="em"><td>每 VR 等值 GW 差額（含新協議）</td>{"".join(f"<td>{n('COST_PropGap_VR_Neg', y, 'sf2')}</td>" for y in (2028, 2029, 2030))}</tr>
<tr><td>　基準（只計已簽合約）</td>{"".join(f"<td>{n('COST_PropGap_VR', y, 'sf2')}</td>" for y in (2028, 2029, 2030))}</tr>
<tr><td>覆蓋率（含新協議）</td>{"".join(f"<td>{n('COST_Coverage_Neg', y, 'pct0')}</td>" for y in (2028, 2029, 2030))}</tr>
<tr><td>累計外部資金需求（含新協議）$B</td>{"".join(f"<td>{n('FND_ExtNeedCumNeg', y, 'f1')}</td>" for y in (2028, 2029, 2030))}</tr>
</tbody></table></div>
<p class="note">期間 {n("Cost!K226", None, "f0")} 年（Assumed，5–10 年）、{n("Cost!K225", None, "yr")} 年起、爬坡沿用 A3 預設。營收不隨之增加，因為基準供給並未卡住需求（容量上限係數＝1）；所以新協議純粹拉高成本、攤薄每 GW 差額，2030 年差額仍為正，累計外部資金需求不變。</p>
{V.src("REV_GrossCapped", "COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExtTrough", "COST_CommitGapCloud", "COST_SpendGap2026Pay", "COST_SpendGap2026Cc", "COST_AdjOIMRep2026Q3", "COST_D22UnderQ3", "COST_NegPay", "CMP_NegGW", "COST_PropGap_VR_Neg", "COST_Coverage_Neg", "FND_ExtNeedCumNeg")}
</details>
</section>""")

    # ── ② 每 VR 等值 GW ──
    # ── ①之三 v0.1.3 A 段更新 ──
    def _rid(sheet, prefix):
        ws = V.loc.wb[sheet]
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and b.startswith(prefix):
                return f"{sheet}!{ws.cell(r, 1).value}"
        raise KeyError(f"{sheet} 找不到列：{prefix}")
    brow = lambda p: _rid("Bridge", p)  # noqa: E731
    erow = lambda p: _rid("EtaDiag", p)  # noqa: E731
    eta_defs = [("計價口徑 f_P（每 GW 合約實付 ÷ TK 持有成本）", "f_P：計價口徑", "Analogy（TK 持有成本；區間＝TK 低／高成本情境）"),
                ("利用率 f_U（TK 實際 ÷ 基準利用率）", "f_U：利用率", "Analogy（TK IF_UtilActual 0.30–0.65；非 Anthropic 實測）"),
                ("工作負載（層級）組合 f_W（TK 參考組合 ÷ 本模型組合）", "f_W：工作負載", "Derived（無區間；Inputs 層級組合區間未接入）"),
                ("輸入輸出組合", None, "無法分離：每任務 token 已含輸入、快取、輸出，並入殘差"),
                ("已解釋部分 f_P × f_U × f_W", "已解釋部分", "Derived"),
                ("未解釋殘差", "未解釋殘差", "Derived（殘差，非獨立估計）")]
    eta_rows = ""
    for lab, key, why in eta_defs:
        if key is None:
            eta_rows += f"<tr><td>{lab}</td><td colspan=\"3\">無法分離</td><td>{why}</td></tr>"
        else:
            rid = erow(key)
            eta_rows += (f"<tr><td>{lab}</td><td>{cell(V, rid + ':D', None, 'f3')}</td><td>{cell(V, rid + ':E', None, 'f3')}</td><td>{cell(V, rid + ':F', None, 'f3')}</td><td>{why}</td></tr>")
    bridge_defs = [("訂閱、席位與 Free 年任務數", "十億任務", "f1"), ("API 年任務數", "十億任務", "f1"), ("付費 token 合計", "T", "f0"), ("API 組合有效單價", "$/M", "f2"),
                   ("營收總額（報導口徑）", "$B", "f1"), ("組合後每 GW 年產能（付費）", "M tok/GW/年", "f0"), ("推論 GW（token 換算）：合計", "GW", "f2"),
                   ("η（有效產出係數", "倍", "f3"), ("有效推論 GW＝", "GW", "f2"), ("研發 GW（殘差）", "GW", "f2"), ("供給 GW 合計", "GW", "f2"),
                   ("供給 VR 等值 GW", "GW", "f2"), ("算力成本（現金口徑）", "$B", "f1"), ("全成本（含股權報酬）", "$B", "f1"),
                   ("每 VR 等值 GW：差額", "$B/GW/年", "sf1")]
    bridge_rows = ""
    for pre, unit, f in bridge_defs:
        rid = brow(pre)
        lab = V.loc.wb["Bridge"].cell(int(re.sub(r"\D", "", V.loc.locate(rid, 2025)[1])), 2).value
        bridge_rows += (f"<tr><td>{html.escape(lab)}</td><td>{unit}</td><td>{cell(V, rid, 2025, f)}</td><td>{cell(V, rid, 2030, f)}</td>"
                        f"<td>{cell(V, rid + ':L', None, f)}</td><td>{cell(V, rid + ':M', None, f)}</td></tr>")
    u2_cols = [("v0.1.2（U2 金額回推）", "v012"), (f"v{ver}（U2 新法）", "base")]
    def u2row(lab, ref, yr, f, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in u2_cols) + "</tr>")
    u2_rows = [u2row(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf2", y == 2030) for y in (2026, 2027, 2028, 2029, 2030)]
    u2_rows += [u2row("覆蓋率 2030", "COST_Coverage", Y, "pct0"), u2row("供給 VR 等值 GW 2030", "CMP_Supply_VReq", Y, "f2"),
                u2row("Nscale 供給 GW 2030", "CMP_Supply_Nscale", Y, "f2"), u2row("Lambda 供給 GW 2030", "CMP_Supply_Lambda", Y, "f2"),
                u2row("Nscale 合約實付 2030 $B", "COST_Pay_Nscale", Y, "f2"), u2row("Lambda 合約實付 2030 $B", "COST_Pay_Lambda", Y, "f2"),
                u2row("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1"), u2row("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f1")]
    lay_rows = "".join(
        f"<tr><td>{L_ZH}</td><td>{n(f'FND_L{k}_GapYear', None, 'yt')}</td><td>{n(f'FND_L{k}_ExtTotal', None, 'f1')}</td><td>{n(f'FND_L{k}_Trough', None, 'f1')}</td></tr>"
        for k, L_ZH in ((1, "① 營運自由現金流自給（不靠任何融資）"), (2, "② 加已到位股權（現行基準）"), (3, "③ 再加條件式與談判中（IPO、Amazon、Google、NVIDIA、AMD；上緣）")))
    exp_rows = "".join(
        f"<tr><td>{lab}</td><td>{n(f'COST_PropGap_VR_Exp{k}', 2029, 'sf2')}</td><td>{n(f'COST_PropGap_VR_Exp{k}', 2030, 'sf2')}</td><td>{n(f'FND_FCFFirstPos_Exp{k}', None, 'yt')}</td>"
        f"<td>{n(f'FND_L1_ExtTotal_Exp{k}', None, 'f1')}</td><td>{n(f'FND_L2_ExtTotal_Exp{k}', None, 'f1')}</td><td>{n(f'FND_L3_ExtTotal_Exp{k}', None, 'f1')}</td><td>{n(f'FND_L2_Trough_Exp{k}', None, 'f1')}</td></tr>"
        for k, lab in (("None", "無需求（變現 0）"), ("Half", "部分變現（0.5）"), ("Full", "完全變現（1）")))
    k14_cols = [("v0.1.3（TK v19.0）", "v013"), ("只換 Luna（IF_TokGW_Luna）", "v14_only_luna"), ("只換其餘 38 名稱", "v14_only_rest"), ("v0.1.4（TK v19.1）", "v014")]
    def k14row(lab, ref, yr, f, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in k14_cols) + "</tr>")
    k14_rows = [k14row("η（2025 校準）", "CMP_Eta", 2025, "f3", True), k14row("供給 VR 等值 GW 2030", "CMP_Supply_VReq", Y, "f3")]
    k14_rows += [k14row(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf3", y == 2030) for y in (2025, 2026, 2027, 2028, 2029, 2030)]
    k14_rows += [k14row("有效推論 GW 2030", "CMP_InfGW_Eff", Y, "f3"), k14row("研發 GW 2030", "CMP_RDGW", Y, "f3"), k14row("覆蓋率 2030（倍）", "COST_Coverage", Y, "f3"), k14row("累計外部資金需求 2026–2030 $B（基準）", "FND_ExtNeedCum", Y, "f1", True),
                 k14row("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2")]
    ev_cols = [("v0.1.1（TK v5.26）", "v011"), ("v0.1.2（v18.0）", "v012"), ("v0.1.3（v19.0）", "v013"), ("v0.1.4（v19.1）", "v014"), ("v0.1.5（v21.2）", "v016"),
               ("v0.1.7–v0.1.8（v21.7）", "v018"), (f"v0.1.9（{tk_ver}）", "base")]
    def evrow(lab, ref, f, em=False):
        ratio = V.get(ref, None, "base") / V.get(ref, None, "v013")
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, None, f, sc)}</td>" for _l, sc in ev_cols) + f"<td>{ratio:.3f}</td></tr>")
    ev_rows = [evrow("η（2025）", "ETA_Eta", "f3", True), evrow("f_P 計價口徑", "ETA_fP", "f3"), evrow("f_U 利用率（不計生產折減）", "ETA_fU", "f3"),
               evrow("f_W 工作負載組合", "ETA_fW", "f3"), evrow("未解釋殘差（不計生產折減；v0.1.3 口徑）", "ETA_Resid", "f3"),
               evrow("f_U' 利用率（計入生產折減）", "ETA_fU_Prod", "f3"), evrow("未解釋殘差（計入生產折減；v0.1.4）", "ETA_Resid2", "f3", True)]
    k15 = [("L1_TokLedger_2026Q2", "token 帳未閉合倍數（模型付費 token ÷ 需求面付費 token）", "f2"),
           ("L1_GapSpendBasis", "TK 落差 2：推論支出計價 ÷ 持有成本（AL_SpendBasis 計價倍數）", "f2"), ("L1_GapSpendBasis_Lo", "　低", "f2"), ("L1_GapSpendBasis_Hi", "　高", "f2"),
           ("L1_GapProduct", "TK 四項乘積（已解釋部分）", "f2"), ("L1_GapProduct_Lo", "　低", "f2"), ("L1_GapProduct_Hi", "　高", "f1"),
           ("L1_GapUnexplained", "TK 落差 5：未解釋", "f2"), ("L1_GapUnexplained_Lo", "　低", "f2"), ("L1_GapUnexplained_Hi", "　高", "f2")]
    k15_html = "".join(f'<tr><td><code>{nm}</code>　{lab}</td><td>{n("TK_" + nm, None, f, "v014")}</td><td>{n("TK_" + nm, None, f)}</td></tr>' for nm, lab, f in k15)
    S.append(f"""<section id="s1c"><h2><span class="no">①之三</span>v0.1.3＋v0.1.4＋v0.1.5 更新了什麼：TK v21.2、v19.1、η 診斷補充；U2 新法、資金三層、持續擴張、η 拆解、逐層對照</h2>
<p class="lead"><b>v0.1.5 摘要</b>：Tokenomics 由 v19.1 升到 {tk_ver}（計價倍數 AL_SpendBasis {n("TK_L1_GapSpendBasis", None, "f2", "v014")} → {n("TK_L1_GapSpendBasis", None, "f2")}）；本模型讀取的 127 個值中只有 10 個變動，全部只進對照列（token 帳未閉合倍數、η 診斷的 TK 並列欄），<b>每 VR 等值 GW 差額、覆蓋率、η、供給 GW、資金與擴張數值全部與 v0.1.4 逐位相同</b>（2030 年差額 {n("COST_PropGap_VR", Y, "sf3")} $B/GW/年；累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0")}；η {n("CMP_Eta2025", None, "f3")}）。token 帳未閉合倍數 {n("Cost!K245", 2026, "f2", "v014")} → {n("Cost!K245", 2026, "f2")}；TK 在 OpenAI 口徑的未解釋落差 L1_GapUnexplained {n("TK_L1_GapUnexplained", None, "f2", "v014")} → {n("TK_L1_GapUnexplained", None, "f2")}，故本模型殘差倒數（計入生產折減）為 TK 的倍數由 {cell(V, erow("比值：本模型殘差倒數"), None, "f3", "v014")} 變為 {cell(V, erow("比值：本模型殘差倒數"), None, "f3")}。</p>
<h3>v0.1.5：TK v19.1 → {tk_ver} 變動的 10 個名稱（只進對照列）</h3>
<div class="tw"><table><thead><tr><th>Tokenomics 名稱</th><th>v0.1.4（v19.1）</th><th>v0.1.5（{tk_ver}）</th></tr></thead><tbody>{k15_html}</tbody></table></div>
<p class="note">這 10 個值只被 Cost 第二十一節（token 帳倍數）、Bridge 與 EtaDiag 第六、七節（TK 並列欄與比值）、Checks 資訊列引用；不被任何命題輸出、資金或判定使用。TK 方向 L1_GapSpendBasis 的計價倍數由 {n("TK_L1_GapSpendBasis", None, "f2", "v014")} 改為 {n("TK_L1_GapSpendBasis", None, "f2")}，是 TK 對 OpenAI 推論支出計價口徑的修正，與本模型自己的計價因子 f_P（取自 Inputs）無關，故 η 不變。</p>
<p class="lead"><b>v0.1.4 摘要</b>：Tokenomics 由 v19.0 升到 v19.1（Luna 的 CSA top-k 1024 → 512；39 個名稱值變動，其中只有 IF_TokGW_Luna 進命題計算），η 由 {n("CMP_Eta2025", None, "f3", "v013")} 微降到 {n("CMP_Eta2025", None, "f3")}，命題輸出（每 VR 等值 GW 差額、覆蓋率、供給 GW、現金流、外部資金需求）與 v0.1.3 <b>逐位相同</b>（2030 年差額 {n("COST_PropGap_VR", Y, "sf3", "v013")} → {n("COST_PropGap_VR", Y, "sf3")} $B/GW/年；累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0")}）：供給 GW 由已簽合約決定，TK 產能變動被 η 的 2025 校準吸收，只移動有效推論 GW 與研發 GW 的分攤（2030 年有效推論 GW {n("CMP_InfGW_Eff", Y, "f3", "v013")} → {n("CMP_InfGW_Eff", Y, "f3")}，研發 GW {n("CMP_RDGW", Y, "f3", "v013")} → {n("CMP_RDGW", Y, "f3")}）。η 診斷補充：利用率因子計入 TK 生產折減（{n("TK_IF_ProdDerate", None, "f2")}）後，未解釋殘差由 {n("ETA_Resid", None, "f3")} 變為 {n("ETA_Resid2", None, "f3")}；η 本身不因診斷而改變。以下兩個表之後，為 v0.1.3 的內容（表中數字為現值，標明 v0.1.3 者為 TK v19.0 當時值）。</p>
<h3>v0.1.4：TK v19.0 → v19.1 的差異拆解（逐群）</h3>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in k14_cols)}</tr></thead><tbody>
{"".join(k14_rows)}
</tbody></table></div>
<p class="note">「只換 Luna」＝只有 IF_TokGW_Luna 取 v19.1、其餘 38 個名稱維持 v19.0；「只換其餘 38」相反。其餘 38 個名稱（IF_FullCost*、IF_Alloc*、L1_RevGW_PQ_VR200* 等）本模型只對照或快照，不進命題輸出；逐名稱新舊值見報告 ①。</p>
<h3>η 與殘差：各 TK 版本（診斷；η 以外的因子為比值，TK 整體產能放大時不動）</h3>
<div class="tw"><table><thead><tr><th>項目</th>{"".join(f"<th>{l}</th>" for l, _s in ev_cols)}<th>v0.1.5 ÷ v0.1.3</th></tr></thead><tbody>
{"".join(ev_rows)}
</tbody></table></div>
<p class="note">殘差＝η ÷（f_P × f_U × f_W）；「計入生產折減」列把 f_U 再乘 TK 生產折減。與 TK 的 OpenAI 口徑 L1_GapUnexplained {cell(V, erow("TK 落差 5"), None, "f2")}（v0.1.3 為 {cell(V, erow("TK 落差 5"), None, "f2", "v013")}）並排：本模型殘差倒數（計入生產折減）{cell(V, erow("殘差倒數（TK 方向；計入"), None, "f2")}，為 TK 的 {cell(V, erow("比值：本模型殘差倒數"), None, "f3")} 倍；兩者口徑不同（Anthropic 對 OpenAI），只作並排。</p>
<p class="lead">v0.1.3 採納外部驗收報告（v0.1.2）的 5 項建議。<b>變動只有三類</b>：①U2 新法改了供給 GW 的取法（本節第一個表）；②新增只供閱讀的對照——資金三層、持續擴張情境（只作情境，不進基準）、η 拆解、Bridge 逐層對照頁；③結論措辭改為條件式。其餘輸出與 v0.1.2 逐格相同。以下先說 U2 新法。U2 新法（chat 端更正公式）：Nscale（{n("SRC_ANT_222", None, "f0")} MW）、Lambda（{n("SRC_ANT_224", None, "f0")} MW）的供給 GW 改取揭露容量，年付款路徑不變（『最高 86』按 MW 比例分攤、依爬坡年度攤提）；因此成本不變、分母（供給 VR 等值 GW）下降，2030 每 VR 等值 GW 差額由 {n("COST_PropGap_VR", Y, "sf2", "v012")} 變為 {n("COST_PropGap_VR", Y, "sf2")}。隱含每 GW 年付款：Nscale {n("COST_U2PayPerGW_Nscale", None, "f1")}、Lambda {n("COST_U2PayPerGW_Lambda", None, "f1")}、兩約合計 {n("COST_U2PayPerGW_Both", None, "f1")} $B/GW/年，高於通用合約價 {n("INP_114", None, "f0")}，表示『最高 86』含揭露 GW 以外的容量或服務，或揭露 MW 低估。v0.1.3 的 Tokenomics 快照改 v19.0，數值與 v18.0 相同（v0.1.4 改 {tk_ver}，見本節開頭）。</p>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in u2_cols)}</tr></thead><tbody>
{"".join(u2_rows)}
</tbody></table></div>
<h3>資金三層（基準）：營運現金流自給 → 已到位股權 → 條件式</h3>
<div class="tw"><table><thead><tr><th>層</th><th>最低現金缺口首次出現年</th><th>累計外部資金需求 2026–2030 $B</th><th>現金谷底 $B</th></tr></thead><tbody>{lay_rows}</tbody></table></div>
<p class="note">第 1 層不計 2025 年底現金（Inputs l1_incl_cash＝0）；第 3 層把 NVIDIA 與 IPO 基石都算入，屬上緣。營運自由現金流首次轉正年：{n("FND_FCFFirstPos", None, "yt")}。</p>
<h3>持續擴張情境（只作情境）：2029、2030 每年追加算力＝當年合約付款的 {n("INP_189", None, "pct0")}</h3>
<div class="tw"><table><thead><tr><th>變現比例</th><th>差額 2029</th><th>差額 2030</th><th>自由現金流首次轉正年</th><th>第 1 層累計外部資金 $B</th><th>第 2 層 $B</th><th>第 3 層 $B</th><th>第 2 層現金谷底 $B</th></tr></thead><tbody>{exp_rows}</tbody></table></div>
<p class="note">基準（只計已簽合約）差額 2029 {n("COST_PropGap_VR", 2029, "sf2")}、2030 {n("COST_PropGap_VR", Y, "sf2")}。追加金額與變現比例為 Assumed（Inputs exp_*）；與 U6 談判中 $150B 並列，不進基準。</p>
<h3>η 拆解（診斷；η 的值不變）</h3>
<p class="note">η（2025）＝token 換算推論 GW ÷ 支出換算推論 GW ＝ {n("ETA_Eta", None, "f3")}。拆成 計價口徑 × 利用率 × 工作負載（層級）組合 × 未解釋殘差；前兩項靠 Tokenomics 的 Analogy 區間，第三項無區間，輸入輸出組合在本模型<b>無法分離</b>（並入殘差）。</p>
<div class="tw"><table><thead><tr><th>因子（η 的方向）</th><th>基準</th><th>低</th><th>高</th><th>依據</th></tr></thead><tbody>
{eta_rows}
</tbody></table></div>
<p class="note">前三項合起來預期 η 約 {n("ETA_Prod", None, "f2")}，實際只有 {n("ETA_Eta", None, "f2")}，差 {n("ETA_Resid", None, "f2")} 倍沒有被解釋（殘差；含輸入輸出組合、每任務 token 倍數、任務數與支出的口徑差等），倒數 {cell(V, erow("殘差倒數"), None, "f2")}（TK 方向；TK 在 OpenAI 口徑的同類項 L1_GapUnexplained {cell(V, erow("TK 落差 5"), None, "f2")}、區間 {cell(V, erow("TK 落差 5") + ":E", None, "f1")}–{cell(V, erow("TK 落差 5") + ":F", None, "f1")}，只作並排）。TK 的 Tokenomics 版本由 v5.26 換到 v18.0 時，η 由 {n("CMP_Eta2025", None, "f3", "v011")} 降到 {n("CMP_Eta2025", None, "f3", "v012")}（見 <code>docs/reports/20261022_v0.1.3.md</code> ④）；v0.1.4 的 {tk_ver} 使 η 再降到 {n("CMP_Eta2025", None, "f3")}（見 <code>{report}</code> ③）。</p>
<h3>需求 → token → GW → 成本逐層對照（Bridge 頁，唯讀）</h3>
<div class="tw"><table><thead><tr><th>層</th><th>單位</th><th>2025</th><th>2030</th><th>2030 低</th><th>2030 高</th></tr></thead><tbody>
{bridge_rows}
</tbody></table></div>
<p class="note">低／高＝2030 敏感度包絡（合約價 8／20、η 沿用／回升、TPU／Trainium／AMD 產出比低／高；Excel Compute 第十四節）。與 Tokenomics 同口徑量（2026）：TK 每 GW 營收 P × Q {cell(V, brow("TK：VR200 每 GW 營收"), 2026, "f1")} vs 本模型 {cell(V, brow("本模型：營收總額 ÷ 付費推論"), 2026, "f1")} $B/GW/年；TK 每 GW 付費 token Q {cell(V, brow("TK：VR200 每 GW 付費 token"), 2026, "f0")} vs 本模型 {cell(V, brow("本模型：每付費推論 VR"), 2026, "f0")} M tok（比值 {cell(V, brow("比值（本模型 ÷ TK）"), 2026, "f2")}）。TK 自己的 token 帳尚未閉合（L1_TokLedger_2026Q2＝{cell(V, brow("TK：token 帳未閉合"), 2026, "f2")} 倍），屬 Tokenomics 待辦，所以這些差額不作為本模型的通過條件。</p>
{V.src("ETA_Eta", "ETA_fP", "ETA_fU", "ETA_fW", "ETA_Resid", "COST_PropGap_VR", "Bridge!" + brow("供給 GW 合計").split("!")[1])}
</section>""")

    # ── ①之四 v0.1.6：正式公開 S-1 ──
    krow = lambda p: _rid("Cost", p)  # noqa: E731
    v16_cols = [("v0.1.5（草擬版 16.3）", "v015"), (f"v{ver}（正式 S-1 16.286）", "base")]
    def v16row(lab, ref, yr, f, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td>' + "".join(f"<td>{n(ref, yr, f, sc)}</td>" for _l, sc in v16_cols) + "</tr>")
    v16_rows = [v16row("2026 營收總額（報導口徑）$B", "REV_GrossCapped", 2026, "f3", True)]
    v16_rows += [v16row(f"每 VR 等值 GW 營收淨額 {y}（$B/GW/年）", "COST_PropRev_VR", y, "f2") for y in (2026, 2027, 2028, 2029, 2030)]
    v16_rows += [v16row(f"每 VR 等值 GW 差額 {y}（$B/GW/年）", "COST_PropGap_VR", y, "sf3", y == 2030) for y in (2026, 2027, 2028, 2029, 2030)]
    v16_rows += [v16row(f"覆蓋率 {y}（倍）", "COST_Coverage", y, "f3") for y in (2026, 2027, 2028, 2029, 2030)]
    v16_rows += [v16row("累計外部資金需求 2026–2030 $B", "FND_ExtNeedCum", Y, "f1", True), v16row("只靠已到位融資的現金谷底 $B", "FND_CashNoExtTrough", None, "f2")]
    s1v = [("2026 上半年營收", "SRC_ANT_414", "SRC_ANT_380", "f3"), ("2026 Q2 營收", "SRC_ANT_420", "SRC_ANT_382", "f3"),
           ("2025 上半年營收", "SRC_ANT_413", "SRC_ANT_381", "f3"), ("2025 全年營收（基準 2025 校準仍用草擬版）", "SRC_ANT_412", "SRC_ANT_001", "f3"),
           ("2026 上半年算力與基礎設施費用", "SRC_ANT_426", "SRC_ANT_394", "f3"), ("2025 全年算力與基礎設施費用", "SRC_ANT_424", "SRC_ANT_252", "f3"),
           ("Google Cloud 合約（至少）", "SRC_ANT_450", "SRC_ANT_397", "f1"), ("Amazon 合約（至少）", "SRC_ANT_451", "SRC_ANT_398", "f1"),
           ("Microsoft Azure 合約（至少）", "SRC_ANT_452", "SRC_ANT_399", "f1"), ("SpaceXAI 合約（最高）", "SRC_ANT_453", "SRC_ANT_229_Hi", "f1"),
           ("Broadcom 租賃", "SRC_ANT_454", "SRC_ANT_401", "f1")]
    s1v_html = "".join(f"<tr><td>{lab}</td><td>{n(a, None, f)}</td><td>{n(b, None, f)}</td></tr>" for lab, a, b, f in s1v)
    s1v_html += (f'<tr><td>Nscale／Lambda 合約</td><td>{n("SRC_ANT_455", None, "f1")}／{n("SRC_ANT_456", None, "f1")}</td><td>兩約合計最高 {n("SRC_ANT_400", None, "f0")}</td></tr>'
                 f'<tr><td>2026-06-30 現金與短期投資</td><td>{n("SRC_ANT_444", None, "f1")}</td><td>轉述未提（v0.1.1 資料缺口 U8）</td></tr>'
                 f'<tr><td>客戶集中</td><td>前五大 {n("SRC_ANT_459", None, "pct0")}（2026 前九個月）；最大揭露客戶 {n("SRC_ANT_461", None, "pct0")}（上半年）</td><td>兩家客戶約 {n("SRC_ANT_405", None, "pct0")}（上半年）</td></tr>'
                 f'<tr><td>淨營收留存率（2026-09-30）</td><td>{n("SRC_ANT_463", None, "pct0")}</td><td>—</td></tr>')
    pg_defs = [("Google Cloud", "合約總額（模型）：Google Cloud", "SRC_ANT_450", "Google Cloud（至少）"), ("Amazon", "合約總額（模型）：Amazon", "SRC_ANT_451", "Amazon（至少）"),
               ("Microsoft Azure", "合約總額（模型）：Microsoft", "SRC_ANT_452", "Microsoft Azure（至少）"), ("SpaceXAI", "合約總額（模型）：SpaceX", "SRC_ANT_453", "SpaceXAI（最高）"),
               ("Broadcom", "合約總額（模型）：Google／Broadcom", "SRC_ANT_454", "Broadcom 租賃"), ("Nscale", "合約總額（模型）：Nscale", "SRC_ANT_455", "Nscale（最高）"),
               ("Lambda", "合約總額（模型）：Lambda", "SRC_ANT_456", "Lambda（約）"), ("AMD（不在承諾表內）", "合約總額（模型）：AMD", "SRC_ANT_457", "AMD（超過")]
    pg_html = "".join(f"<tr><td>{lab}</td><td>{n(krow(m), 2025, 'f1')}</td><td>{n(src, None, 'f1')}</td><td>{n(krow('(d) 模型 ÷ 正式 S-1 − 1：' + g), 2025, 'pct0')}</td></tr>"
                      for lab, m, src, g in pg_defs)
    h_rows = "".join(f'<tr><td>{lab}</td>' + "".join(f"<td>{n(ref, y, f)}</td>" for y in (2026, 2027, 2028, 2029, 2030)) + "</tr>"
                     for lab, ref, f in (("η（B 情境：2026 起＝η_2026H1）", krow("(c) B 情境 η 路徑"), "f3"), ("η（基準）", "CMP_Eta", "f3"),
                                         ("容量上限係數（B 情境）", "CMP_CapFactor_EtaH1", "f2"), ("每 VR 等值 GW 營收淨額（B 情境）", "COST_PropRev_VR_EtaH1", "f2"),
                                         ("每 VR 等值 GW 差額（B 情境）", "COST_PropGap_VR_EtaH1", "sf3"), ("差：B 情境 − 基準", krow("(c) 差：B 情境 − 基準（每 VR"), "sf3"),
                                         ("覆蓋率（B 情境）", "COST_Coverage_EtaH1", "f3"), ("研發 GW（B 情境）", krow("(c) B 情境 研發 GW"), "f2"), ("研發 GW（基準）", "CMP_RDGW", "f2")))
    S.append(f"""<section id="s1d"><h2><span class="no">①之四</span>v0.1.6 更新了什麼：正式公開 S-1（2026-10-26，SEC EDGAR）</h2>
<p class="lead"><b>正式 S-1 沒有翻轉任何結論。</b>基準只改一處（工作單 W2）：2026 年半校準用的上半年實際營收，由草擬版轉述的 {n("SRC_ANT_380", None, "f1")} 改為正式 S-1 的 {n("SRC_ANT_414", None, "f3")}（差 {n("Checks!" + chk_of("f1_h1"), None, "f3")}，{n("Checks!" + chk_of("f1_h1pct"), None, "pct1")}；工作單 5% 門檻內），2026 營收 {n("REV_GrossCapped", 2026, "f2", "v015")} → {n("REV_GrossCapped", 2026, "f3")}；2026 每 VR 等值 GW 營收 {n("COST_PropRev_VR", 2026, "f2", "v015")} → {n("COST_PropRev_VR", 2026, "f2")}，差額 2026–2028 仍為負、2029 起轉正（2030 {n("COST_PropGap_VR", Y, "sf3", "v015")} → {n("COST_PropGap_VR", Y, "sf3")}），2030 覆蓋率 {n("COST_Coverage", Y, "pct0")}，累計外部資金需求仍為 {n("FND_ExtNeedCum", Y, "f0")}。正式 S-1 的其他數字（季度、營收成本、承諾、逐家金額、集中度、現金）只進對照列（Excel Cost 第二十四節），不回饋基準；η 改用 2026 上半年的值只作情境（見下），基準 η 不改。</p>
<div class="tw"><table><thead><tr><th>輸出</th>{"".join(f"<th>{l}</th>" for l, _s in v16_cols)}</tr></thead><tbody>
{"".join(v16_rows)}
</tbody></table></div>
<h3>正式 S-1 vs 草擬版轉述（$B；SRC_ANT_410–464，草擬版舊列值不改）</h3>
<div class="tw"><table><thead><tr><th>項目</th><th>正式 S-1</th><th>草擬版轉述（v0.1.1–v0.1.5）</th></tr></thead><tbody>{s1v_html}</tbody></table></div>
<p class="note">年度與上半年營收、算力與基礎設施費用、營業損益、6-30 現金與 NRR 直接讀自 S-1（Verified；期中數字未經查核）。季度數、營收成本（Note 8）、營業現金流、承諾三段、逐家金額與集中度：sec.gov 擋自動工具、可讀取的 S-1 文字只到風險因素，改取二手轉載（Kepler、sec-api.io、Tunguz、AlphaSense 等），標 Interested-party，列為待核對缺口。S-1 未揭露 2026 Q3 營收與承諾總額（留白），不填推論值。</p>
<h3>(a) 算力與基礎設施費用 vs 模型 2026 算力成本</h3>
<p class="note">S-1 上半年費用 {n(krow("(a) S-1：2026 上半年算力"), 2026, "f3")}；模型 2026 合約實付折半 {n(krow("(a) 模型：2026 合約實付"), 2026, "f2")}（模型 ÷ S-1 − 1＝{n("COST_S1fGapPayH1", None, "pct0")}），含自建資本支出折半 {n(krow("(a) 模型：2026 算力成本"), 2026, "f2")}（{n("COST_S1fGapCcH1", None, "pct0")}）。口徑不同：S-1 是損益費用（預付雲端款先列資產、融資租賃以折舊與利息入帳），模型是現金。S-1 上半年資本支出年化只有 {n(krow("(a) S-1：2026 上半年資本支出"), 2026, "f2")}，模型 2026 自建資本支出（Fluidstack $50B 三年均攤）是它的 {n(krow("(a) 模型：2026 自建資本支出"), 2026, "f1")} 倍——自建支出在 S-1 現金流中看不到，可能由 Fluidstack 出資後以租賃回租，或時點較晚（待 Andy 判斷，基準不改）。</p>
<h3>(b) 營收成本 vs 模型 2025 推論算力</h3>
<p class="note">S-1 2025 營收成本 {n(krow("(b) S-1：2025 年營收成本"), 2025, "f3")}，模型 2025 推論算力（說明書算力 − 訓練）{n(krow("(b) 模型：2025 推論算力成本"), 2025, "f2")}（S-1 ÷ 模型 − 1＝{n("COST_S1fGapCOR2025", None, "pct0")}）；營收成本占算力與基礎設施費用 2025 {n(krow("(b) 營收成本占算力"), 2025, "pct0")}、2026 上半年 {n(krow("(b) 營收成本占算力"), 2026, "pct0")}，模型推論占算力 2025 {n(krow("(b) 模型：推論算力占"), 2025, "pct0")}、2026 {n(krow("(b) 模型：推論算力占"), 2026, "pct0")}。</p>
<h3>(c) η 診斷：2026 上半年 η 與「η＝2026H1」B 情境（只作情境）</h3>
<p class="note">以 2026 上半年營收成本年化 {n(krow("(c) S-1：2026 上半年營收成本"), 2026, "f2")} ÷ 既有雲端加權合約價 {n(krow("(c) 既有雲端加權合約價"), 2026, "f2")}（與 2025 η 同法）得上半年推論 {n(krow("(c) 2026 上半年推論 GW（支出換算）"), 2026, "f3")} GW；模型 2026 token 換算 GW 依上半年營收占比 {n(krow("(c) 上半年占 2026 營收"), 2026, "pct0")} 折成上半年平均 {n(krow("(c) 2026 上半年推論 GW（token"), 2026, "f3")} GW → <b>η_2026H1＝{n("COST_Eta2026H1", None, "f3")}</b>，為基準 η（2025）{n("CMP_Eta2025", None, "f3")} 的 {n(krow("(c) η_2026H1 ÷ 基準"), 2026, "f2")} 倍（不折上半年的上緣 {n(krow("(c) 替代：全年 token"), 2026, "f3")}；若 2025 也改用 S-1 營收成本，η（2025）為 {n("COST_Eta2025COR", None, "f3")}）。η 較高表示每 GW 兌現的 token 較多；代入後有效推論 GW 變少、研發 GW 變多，已簽合約的供給並未被需求卡住（容量上限係數仍為 1），<b>所以每 VR 等值 GW 營收、差額、覆蓋率與基準完全相同</b>。</p>
<div class="tw"><table><thead><tr><th>B 情境（η＝2026H1）</th><th>2026</th><th>2027</th><th>2028</th><th>2029</th><th>2030</th></tr></thead><tbody>{h_rows}</tbody></table></div>
<h3>(d) S-1 承諾三段 vs 模型逐家合約總額（Cost 第十一、十二節口徑）</h3>
<p class="note">S-1：6-30 尚未起租租賃 {n("SRC_ANT_445", None, "f1")}＋6-30 固定或最低採購承諾 {n("SRC_ANT_446", None, "f1")}＋6-30 後新簽與修訂 {n("SRC_ANT_447", None, "f0")}＝{n(krow("(d) S-1 三段合計"), 2025, "f1")}（不含已起租租賃與短期承諾；S-1 總額留白）。模型全部新合約 {n(krow("(d) 模型：全部新合約加總"), 2025, "f1")}（模型 ÷ S-1 − 1＝{n("COST_S1fCommitGap", None, "pct0")}）；七家具名合約 S-1 {n(krow("(d) S-1：七家具名合約合計"), 2025, "f1")} vs 模型 {n(krow("(d) 模型：同七家合約總額"), 2025, "f1")}（{n("COST_S1fNamedGap", None, "pct0")}）。模型只計已簽且已揭露金額的合約，6-30 後新簽的約 245 大多不在模型內；只對照，不改供給。</p>
<div class="tw"><table><thead><tr><th>合約</th><th>模型合約總額</th><th>正式 S-1</th><th>模型 ÷ S-1 − 1</th></tr></thead><tbody>{pg_html}</tbody></table></div>
{V.src("REV_GrossCapped", "COST_PropRev_VR", "COST_Eta2026H1", "COST_PropGap_VR_EtaH1", "COST_S1fCommitGap", "SRC_ANT_414")}
</section>""")

    V.default_sc = "base"
    _hist = S[_hist_start:]
    del S[_hist_start:]

    S.append(f"""<section id="s2"><h2><span class="no">②</span>每 VR 等值 GW：營收 vs 全成本（2025–2030）</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>營收淨額（截頂後；扣雲端平台抽成）</span><span><i style="background:var(--cost)"></i>全成本（算力＋非算力＋股權報酬，現金口徑）</span><span><b style="color:var(--gap)">紅</b>／<b style="color:var(--pos)">綠</b>字＝差額</span><span>2025 長條截斷，白字為實值</span></div>
<div class="cw">{chart_gw(V)}</div>
{year_table(V, [("營收淨額", "COST_PropRev_VR", "f1", ""), ("　算力成本（合約實付＋自建資本支出）", "COST_PropCompute_VR", "f1", ""),
                ("　非算力成本（不含股權報酬）", "COST_PropNonComp_VR", "f1", ""), ("　股權報酬", "COST_PropSBC_VR", "f1", ""),
                ("全成本", "COST_PropFull_VR", "f1", ""), ("差額（命題 1；現金・每 VR 等值 GW）", "COST_PropGap_VR", "sf1", "em"),
                ("覆蓋率", "COST_Coverage", "pct0", ""), ("分母：供給 VR 等值 GW", "CMP_Supply_VReq", "f2", ""),
                ("對照：經濟口徑差額（自建改以持有成本年化）", "COST_PropGapEcon_VR", "sf1", "sep"), ("對照：每實體 GW 差額", "COST_PropGap_GW", "sf1", ""),
                ("分母：供給實體 GW", "CMP_SupplyGW", "f2", ""), ("差額總額（$B）", "COST_GapCash", "sf1", "")], head="$B/GW/年")}
<p class="note">2025 分母只有 {n("CMP_Supply_VReq", 2025, "f2")} VR 等值 GW（2025 機隊為 Trainium2、TPU v6e、Hopper，換成 VR200 等值只有實體 {n("CMP_SupplyGW", 2025, "f2")} GW 的 6%），每 GW 數字因此偏大，請看每實體 GW 列（{n("COST_PropGap_GW", 2025, "sf1")}）。三種口徑的正負一致：2026–2029 為負（2026 年 v0.1.8 起因 S-1/A Q3 股權報酬與非現金捐贈計入全成本、未歸屬餘額 2026Q4 起付而轉負）、2030 轉正；v0.1.7 起自建資本支出為 0（W2），現金口徑與經濟口徑相同。2027 年起的成本跳升主要來自 S-1/A 承諾（含未歸屬餘額 2028 年 {n("COST_Pay_S1Unalloc", 2028, "f1")}，依 S-1/A 付款表結構）與 2026 年起以 S-1／S-1/A 校準的非算力成本。</p>
{V.src("COST_PropRev_VR", "COST_PropCompute_VR", "COST_PropNonComp_VR", "COST_PropSBC_VR", "COST_PropFull_VR", "COST_PropGap_VR", "COST_Coverage", "CMP_Supply_VReq", "COST_PropGapEcon_VR", "COST_PropGap_GW", "CMP_SupplyGW", "COST_GapCash")}
</section>""")

    # ── ③ 營收結構與 run-rate ──
    rr_rows = []
    for sid in RUNRATE_IDS:
        _, lab = period_pos(V.loc, sid)
        rr_rows.append(f"<tr><td>{lab}</td><td>{n(f'SRC_ANT!{sid}:E', None, 'f0')}</td><td>{html.escape(sid)}</td></tr>")
    rr_rows.append(f"<tr><td>{period_pos(V.loc, RUNRATE_RANGE_ID)[1]}（投資人預期）</td><td>{n(f'SRC_ANT!{RUNRATE_RANGE_ID}:F', None, 'f0')}–{n(f'SRC_ANT!{RUNRATE_RANGE_ID}:G', None, 'f0')}</td><td>{RUNRATE_RANGE_ID}</td></tr>")
    S.append(f"""<section id="s3"><h2><span class="no">③</span>營收結構、雲端平台抽成、年化營收里程碑</h2>
<div class="legend"><span><i style="background:var(--sub)"></i>個人訂閱（Pro、Max）</span><span><i style="background:var(--seat)"></i>企業席位（Team、Enterprise）</span><span><i style="background:var(--api)"></i>API（含經雲端市集）</span><span><i style="background:var(--ads)"></i>其他</span><span><i style="background:var(--net);height:3px"></i>淨額（扣雲端平台抽成）</span><span><i style="background:var(--tgt);transform:rotate(45deg)"></i>管理層目標（舊版）</span></div>
<div class="cw">{chart_rev(V)}</div>
{year_table(V, [("個人訂閱（Pro、Max）", "REV_SubConsumer", "f1", ""), ("企業席位訂閱（Team、Enterprise）", "REV_SubSeats", "f1", ""), ("API", "REV_API", "f1", ""), ("其他", "REV_Other", "f2", ""),
                ("總額（報導口徑；截頂後）", "REV_GrossCapped", "f1", "em"), ("雲端平台抽成（截頂後）", "REV_PartnerShareCapped", "f1", ""), ("淨額（命題與現金流用）", "REV_NetCapped", "f1", "em"),
                ("企業占（個人＋企業）", "Revenue!R60", "pct0", "sep"), ("　其中 API 經雲端通路的營收", "REV_Channel", "f1", ""),
                ("管理層目標（只對照）", "RVS_Target", "f1", ""), ("差距（目標 − 正向總額）", "RVS_Gap", "f1", "")])}
<p class="note">報導營收為總額（AWS Bedrock、Google Vertex 等雲端市集以總額認列），本模型扣掉平台抽成（2026 起 {n("Revenue!R55", 2026, "pct0")}）後的淨額才進命題與現金流；2030 年抽成 {n("REV_PartnerShareCapped", Y, "f1")}。廣告＝0（D3）。容量上限係數各年皆為 {n("CMP_CapFactor", Y, "f0")}（未受限），所以各線未截頂值＝截頂值。管理層目標都早於 2026 年的實際爆發，全部低於模型正向值。</p>
<h3>模型年營收 vs 年化營收（run-rate）里程碑（D23：只對照，不校準）</h3>
<div class="legend"><span><i style="background:var(--rev)"></i>模型年營收總額（年平均水準）</span><span><i style="background:var(--cost);border-radius:50%"></i>報導年化營收</span><span><i style="background:var(--tgt)"></i>2026 年底投資人預期區間</span></div>
<div class="cw">{chart_runrate(V)}</div>
<div class="tw"><table><thead><tr><th>對照</th><th>2025</th><th>2026</th></tr></thead><tbody>
<tr class="em"><td>模型年營收總額</td><td>{n("REV_GrossCapped", 2025)}</td><td>{n("REV_GrossCapped", 2026)}</td></tr>
<tr><td>由 run-rate 推估的年營收（2025：年初、年底年化的對數平均；2026：上半年實際＋下半年 65 → 年底預期中點）</td><td>{n("Revenue!R72", 2025)}</td><td>{n("REV_RREst2026", None)}</td></tr>
<tr><td>模型 ÷ 推估 − 1</td><td>{n("Revenue!R73", 2025, "pct0")}</td><td>{n("REV_RRGap2026", None, "pct0")}</td></tr>
</tbody></table></div>
<details><summary class="note">年化營收里程碑明細（SRC_ANT；公司新聞稿與媒體轉述，多為利害關係方說法，含第三方雲端通路總額）</summary>
<div class="tw"><table><thead><tr><th>月份</th><th>年化營收 $B</th><th>來源列</th></tr></thead><tbody>{"".join(rr_rows)}</tbody></table></div></details>
<p class="note">2025 模型以認列營收 {n("SRC_ANT_001", None, "f2")} 校準，高於 run-rate 推估 {n("Revenue!R73", 2025, "pct0")}；2026 模型（v0.1.1 U1；v0.1.6 起上半年改正式 S-1、v0.1.8 起 Q3 改 S-1/A 實績）為「上半年實際 {n("SRC_ANT_414", None, "f3")}＋Q3 實際 {n("SRC_ANT_488", None, "f3")}＋Q4 持平於 Q3」＝{n("REV_GrossCapped", 2026)}，低於里程碑推估 {n("REV_RRGap2026", None, "pct0")}（若年底真到 100–120，2026 會被低估；年底年化的報導互相矛盾：分析師 110、Euronews「低於 100」）。OpenAI 內部備忘錄估 Anthropic 2026-04 年化以淨額計只有 {n("SRC_ANT!SRC_ANT_035:E", None, "f0")}（對公司說的 {n("SRC_ANT!SRC_ANT_021:E", None, "f0")}），差在雲端總額認列。</p>
{V.src("REV_SubConsumer", "REV_SubSeats", "REV_API", "REV_Other", "REV_GrossCapped", "REV_PartnerShareCapped", "REV_NetCapped", "Revenue!R60", "REV_Channel", "RVS_Target", "RVS_Gap", "Revenue!R55", "Revenue!R72", "Revenue!R73", "REV_RREst2026", "REV_RRGap2026", "SRC_ANT_001")}
</section>""")

    # ── ④ 算力 ──
    contracts = [  # (標籤, 供給 GW 具名, 實付具名, 模型合約總額列, 對照金額列, 差距列)；對照金額＝草擬 S-1（WSJ 轉述）；SpaceX、AMD 沿用招股書列
        ("Google Cloud TPU", "CMP_Supply_Google", "COST_Pay_Google", "Cost!K121", "SRC_ANT!SRC_ANT_397:E", "Cost!K143"),
        ("Google／Broadcom 次世代 TPU", "CMP_Supply_Broadcom", "COST_Pay_Broadcom", "Cost!K122", "SRC_ANT!SRC_ANT_401:F", "Cost!K146"),
        ("Amazon（Trainium）", "CMP_Supply_Aws", "COST_Pay_Aws", "Cost!K123", "SRC_ANT!SRC_ANT_398:E", "Cost!K144"),
        ("Microsoft Azure＋NVIDIA", "CMP_Supply_Azure", "COST_Pay_Azure", "Cost!K124", "SRC_ANT!SRC_ANT_399:E", "Cost!K145"),
        ("AMD MI455X（招股書）", "CMP_Supply_Amd", "COST_Pay_Amd", "Cost!K125", "Cost!K138", "Cost!K148"),
        ("SpaceX／xAI Colossus（招股書）", "CMP_Supply_Spacex", "COST_Pay_Spacex", "Cost!K126", "Cost!K137", "Cost!K147"),
        ("Nscale（兩約合計 86 按 MW 分攤）", "CMP_Supply_Nscale", "COST_Pay_Nscale", "Cost!K127", "Cost!K184", "Cost!K186"),
        ("Lambda（兩約合計 86 按 MW 分攤）", "CMP_Supply_Lambda", "COST_Pay_Lambda", "Cost!K128", "Cost!K185", "Cost!K188"),
        ("Volta Infra", "CMP_Supply_Volta", "COST_Pay_Volta", "Cost!K129", None, None),
        ("Akamai", "CMP_Supply_Akamai", "COST_Pay_Akamai", "Cost!K130", None, None),
    ]
    crow = []
    for lab, g, p, tot, pros, gap in contracts:
        gg = V.get(gap) if gap else 0
        warn = ' class="neg"' if gap and abs(gg) > V.get("INP_179") else ""
        crow.append(f"<tr><td>{lab}</td><td>{cell(V, g, 2027, 'f2')}</td><td>{cell(V, g, 2030, 'f2')}</td><td>{cell(V, p, 2028)}</td><td>{cell(V, p, 2030)}</td>"
                    f"<td>{cell(V, tot)}</td><td>{cell(V, pros) if pros else '—'}</td><td{warn}>{cell(V, gap, None, 'pct0') if gap else '—'}</td></tr>")
    crow.append(f"<tr><td>既有雲端（2025 前簽）</td><td>{cell(V, 'CMP_Supply_Legacy', 2027, 'f2')}</td><td>{cell(V, 'CMP_Supply_Legacy', 2030, 'f2')}</td><td>{cell(V, 'COST_Pay_Legacy', 2028)}</td><td>{cell(V, 'COST_Pay_Legacy', 2030)}</td><td>—</td><td>—</td><td>—</td></tr>")
    crow.append(f"<tr><td>S-1/A 承諾未歸屬餘額（v0.1.7；v0.1.8 起三段 871.4、付款依 S-1/A 結構；通用合約價 → VR200 族）</td><td>{cell(V, 'CMP_Supply_S1Unalloc', 2027, 'f2')}</td><td>{cell(V, 'CMP_Supply_S1Unalloc', 2030, 'f2')}</td><td>{cell(V, 'COST_Pay_S1Unalloc', 2028)}</td><td>{cell(V, 'COST_Pay_S1Unalloc', 2030)}</td><td>{cell(V, 'COST_S1Unalloc')}</td><td>—</td><td>—</td></tr>")
    crow.append(f"<tr><td>自建（v0.1.7 起 GW 法；Wattonomics 自建 GW 0）</td><td>{cell(V, 'CMP_Supply_Owned', 2027, 'f2')}</td><td>{cell(V, 'CMP_Supply_Owned', 2030, 'f2')}</td><td>{cell(V, 'COST_OwnedCapex', 2028)}</td><td>{cell(V, 'COST_OwnedCapex', 2030)}</td><td>—</td><td>—</td><td>—</td></tr>")
    crow.append(f"<tr class=\"em\"><td>合計</td><td>{cell(V, 'CMP_SupplyGW', 2027, 'f2')}</td><td>{cell(V, 'CMP_SupplyGW', 2030, 'f2')}</td><td>{cell(V, 'COST_Compute', 2028)}</td><td>{cell(V, 'COST_Compute', 2030)}</td>"
                f"<td>{cell(V, 'COST_CommitAllS1')}</td><td>—</td><td>—</td></tr>")
    S.append(f"""<section id="s4"><h2><span class="no">④</span>算力：總需求 vs 已簽合約＋自建供給（GW，IT 電力）</h2>
<div class="legend"><span><i style="background:var(--inf)"></i>有效推論 GW</span><span><i style="background:var(--rd)"></i>研發 GW（殘差）</span><span><i style="background:var(--con)"></i>合約供給</span><span><i style="background:var(--own)"></i>自建供給</span><span>每年左＝需求、右＝供給</span></div>
<div class="cw">{chart_compute(V)}</div>
{year_table(V, [("有效推論 GW", "CMP_InfGW_Eff", "f2", ""), ("研發 GW（殘差）", "CMP_RDGW", "f2", ""), ("總需求 GW", "CMP_DemandGW", "f2", "em"),
                ("合約供給 GW", "CMP_SupplyContractGW", "f2", ""), ("自建供給 GW", "CMP_Supply_Owned", "f2", ""), ("供給 GW 合計", "CMP_SupplyGW", "f2", "em"),
                ("供給 VR 等值 GW（命題分母）", "CMP_Supply_VReq", "f2", ""), ("推論可用 GW（容量上限）", "CMP_InfAvailGW", "f2", "")], head="GW")}
<p><b>供給 2028 年後遞減</b>：模型只計已簽合約與 S-1/A 承諾（未歸屬餘額依 S-1/A 付款表結構，2028 年 {n("CMP_Supply_S1Unalloc", 2028, "f2")} GW），SpaceX 月費 2029-05 到期；v0.1.7 起自建（Fluidstack）不另計。供給 GW 由 2028 年 {n("CMP_SupplyGW", 2028, "f1")} 降到 2030 年 {n("CMP_SupplyGW", Y, "f1")}；這是 2030 年轉正的主因之一。研發 GW 是殘差（供給扣閒置後減推論），2027 年 {n("CMP_RDGW", 2027, "f1")} GW，比 Tokenomics 研發錨點高兩個數量級（Tokenomics 缺口）。</p>
<h3>逐家合約與草擬 S-1 對帳（U3；v0.1.7 起模型合約總額改讀正式 S-1，對照欄仍為草擬版轉述，正式 S-1 對照見 ①之五）</h3>
<div class="tw"><table><thead><tr><th>合作方</th><th>GW 2027</th><th>GW 2030</th><th>實付 2028</th><th>實付 2030</th><th>模型合約總額</th><th>S-1 草稿金額</th><th>模型 ÷ S-1 − 1</th></tr></thead><tbody>
{"".join(crow)}
</tbody></table></div>
<p class="note">基準以「GW × 每 GW 年合約價 {n("INP_114", None, "f0")} × 持有比」或揭露金額計（D10）；逐家差距超過 {n("INP_179", None, "pct0")} 標紅並觸發 Checks WARN（AMD 模型 {n("Cost!K125", None, "f1")} 對招股書 {n("Cost!K138", None, "f0")}，SpaceX 錨在月費 $45B 對招股書「最高」{n("Cost!K137", None, "f1")}，Broadcom 模型 {n("Cost!K122", None, "f1")} 對 S-1 草稿「至少」{n("SRC_ANT!SRC_ANT_401:F", None, "f0")}）。合計對得上是誤差互相抵銷。三個口徑合計（模型 ÷ 報導 − 1）：雲端付款義務約 {n("Cost!K191", None, "f0")} 為 {n("COST_CommitGapCloud", None, "pct1")}、已鎖定 >{n("Cost!K194", None, "f0")} 為 {n("COST_CommitGapLocked", None, "pct1")}、含談判中約 {n("Cost!K197", None, "f0")} 為 {n("COST_CommitGapAll", None, "pct1")}。逐家錨定金額的敏感度（S4）：2026 差額 {n("COST_PropGap_VR_Anchor", 2026, "sf1")}、2030 {n("COST_PropGap_VR_Anchor", Y, "sf1")}，仍不需外部資金（見 ⑤）。機房租約 {n("Compute!C170", None, "f2")} GW 只列對照、不計供給。</p>
{V.src("CMP_InfGW_Eff", "CMP_RDGW", "CMP_DemandGW", "CMP_SupplyContractGW", "CMP_Supply_Owned", "CMP_SupplyGW", "CMP_Supply_VReq", "CMP_InfAvailGW", "Cost!K121", "Cost!K133", "Cost!K141", "Cost!K143", "COST_PropGap_VR_Anchor", "Compute!C170", "INP_114", "INP_179", "COST_CommitGapCloud", "Cost!K184")}
</section>""")

    # ── ⑤ 現金與融資 ──
    scen_rows = [("基準（已到位＋IPO＋私募淨額，v0.2.0 B1）", "FND_CashEnd", "FND_ExtNeedCum"), ("S1 ＋IPO（v0.2.0：已進基準，與基準相同）", "Funding!F49", "Funding!F50"),
                 ("S2 ＋條件式（Amazon、Google 剩餘、AMD 剩餘；NVIDIA 已併入私募）", "Funding!F54", "Funding!F55"), ("S3 ＋IPO＋條件式", "Funding!F59", "Funding!F60"),
                 ("S4 逐家錨定招股書金額", "Funding!F64", "Funding!F65"), ("S5 加計機房租約租金", "Funding!F69", "Funding!F70"),
                 ("S6 D22 舊法低估再補足（v0.1.7 基準已以 S-1 校準，本列為再加碼）", "Funding!F74", "Funding!F75"), ("S7 S4＋S5＋S6 合併", "Funding!F79", "Funding!F80")]
    sr = []
    for lab, cash, cum in scen_rows:
        sr.append(f'<tr class="{"em" if cash == "FND_CashEnd" else ""}"><td>{lab}</td>' + "".join(f"<td>{cell(V, cash, y, 'f1')}</td>" for y in YEARS[1:])
                  + f"<td>{cell(V, cum, Y, 'f1')}</td></tr>")
    rt = []
    for rid, lab in (("F88", "Amazon"), ("F89", "Google"), ("F90", "Broadcom（債務）"), ("F91", "Microsoft"), ("F92", "NVIDIA（間接）"), ("F93", "AMD"), ("F94", "合計")):
        rt.append(f'<tr class="{"em" if rid == "F94" else ""}"><td>{lab}</td>' + "".join(
            f"<td>{cell(V, f'Funding!{rid}:{c}', None, 'pct0' if c in 'HI' else 'f1')}</td>" for c in "DEFGHI") + "</tr>")
    S.append(f"""<section id="s5"><h2><span class="no">⑤</span>現金路徑與融資</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>年底現金（基準）</span><span><i style="background:var(--cost)"></i>年底現金（S7：錨定招股書＋機房租金＋非算力補足）</span><span><i style="background:var(--gap);height:3px"></i>最低現金（次年非算力 6 個月）</span></div>
<div class="cw">{chart_cash(V)}</div>
{year_table(V, [("營收淨額", "Funding!F01", "f1", ""), ("算力成本（現金）", "Funding!F02", "f1", ""), ("　其中自建資本支出", "Funding!F03", "f1", ""),
                ("非算力成本（不含股權報酬）", "Funding!F04", "f1", ""), ("自由現金流（股權報酬加回）", "FND_FCF", "sf1", "em"),
                ("期初現金（①）", "FND_CashOpen", "f1", ""), ("已到位融資（②）", "FND_Committed", "f1", ""), ("當年外部資金需求（③）", "FND_ExtNeed", "f1", "em"),
                ("年底現金", "FND_CashEnd", "f1", "em"), ("最低現金", "FND_MinCash", "f1", ""), ("現金餘裕（年底 − 最低）", "FND_Headroom", "f1", ""),
                ("累計外部資金需求（命題 2）", "FND_ExtNeedCum", "f1", "em")])}
<h3>已到位融資（2026；來源順序 ②）</h3>
<div class="tw"><table><thead><tr><th>來源</th><th>$B</th></tr></thead><tbody>
<tr><td>Series G-1／G-2（2026 Q1；正式 S-1 Note 10；GIC、Coatue 領投）</td><td>{n("Funding!F12", 2026, "f1")}</td></tr>
<tr><td>Amazon Series G-2（2026 Q2；19.3 百萬股；Amazon 10-Q 與 S-1 Note 4、10）</td><td>{n("Funding!F13", 2026, "f1")}</td></tr>
<tr><td>Google Series G-2（2026 Q2；38.6 百萬股 × 259.1364；S-1 Note 4、10：與 Series H 分開，v0.1.7 未計）</td><td>{n("Funding!F17", 2026, "f1")}</td></tr>
<tr><td>Series H-1／H-2（2026 Q2；S-1 Note 10；含 Google 10、Amazon 5：{n("Funding!F15", 2026, "f0")}）</td><td>{n("Funding!F14", 2026, "f1")}</td></tr>
<tr><td>Microsoft $5B（2025-11，已在 2025 年底現金內）</td><td>{n("Funding!F16", 2026, "f0")}</td></tr>
<tr class="em"><td>合計（對照：S-1 現金流量表特別股發行收入 {n("SRC_ANT_472", None, "f3")}，差 {n("FND_CommittedS1Gap", None, "sf2")}）</td><td>{n("Funding!F18", 2026, "f1")}</td></tr>
</tbody></table></div>
<h3>條件式、IPO 與敏感度情境（同一規則；年底現金與 2030 累計外部資金需求）</h3>
<div class="tw"><table><thead><tr><th>情境</th><th>2026</th><th>2027</th><th>2028</th><th>2029</th><th>2030</th><th>累計外部資金</th></tr></thead><tbody>
{"".join(sr)}
</tbody></table></div>
<p class="note">條件式（未到位）：Amazon {n("Funding!F22", 2027, "f0")}（依算力交付）、Google {n("Funding!F23", 2027, "f1")}（初步招股說明書：私募完成後 2026 NPA 剩餘；另 {n("SRC_ANT_526", None, "f1")} 轉入同步私募）、NVIDIA {n("Funding!F24", 2026, "f1")}（併入同步私募 {n("SRC_ANT_527", None, "f1")}）、AMD 認股剩餘 {n("Funding!F25", 2027, "f1")}（另 {n("SRC_ANT_528", None, "f1")} 在私募）；以上條件式不計入基準。IPO＋同步私募淨額 {n("FND_IPONet", None, "f1")}（淨募資 − RSU 代扣稅款；Microsoft 私募 {n("SRC_ANT_531", None, "f1")}、合約下限 {n("SRC_ANT_573", None, "f1")}）已依 B1 計入基準。<b>或有負債（只列示）</b>：A&R 循環信用額度 {n("Funding!F81", None, "f0")}（2031-09 到期；最低流動性條款 {n("SRC_ANT_486", None, "f0")}，自 2026-12-31 起，低於模型最低現金）、Broadcom 租賃融資最高 {n("Funding!F82", None, "f0")}、表外 Google TPU 融資 SPV {n("Funding!F83", None, "f0")}（單一來源，未證實），合計 {n("FND_Contingent", None, "f0")}；合約 2030 年後剩餘承諾 {n("FND_CommitAfter2030", None, "f1")}（招股書稱約 {n("Funding!F86", None, "pct0")} 不可取消）。</p>
<h3>回流對照（D18：只對照、不沖銷）</h3>
<div class="tw"><table><thead><tr><th>投資人兼算力供應商</th><th>已到位投資</th><th>條件式（未到位）</th><th>合約總額（招股書）</th><th>模型 2026–30 付款</th><th>已到位 ÷ 付款</th><th>（到位＋條件）÷ 合約</th></tr></thead><tbody>
{"".join(rt)}
</tbody></table></div>
<p class="note">策略投資人已投入的錢只相當於 2026–2030 付給它們的算力款的 {n("Funding!F94:H", None, "pct0")}：算力錢主要不是由賣算力的一方回流，而是由 2026 年的財務投資人與 2030 年起的營收提供。</p>
{V.src("Funding!F01", "FND_FCF", "FND_CashOpen", "FND_Committed", "FND_ExtNeed", "FND_CashEnd", "FND_MinCash", "FND_Headroom", "FND_ExtNeedCum", "Funding!F12", "Funding!F21", "Funding!F49", "Funding!F79", "FND_Contingent", "FND_CommitAfter2030", "Funding!F88")}
</section>""")

    # ── ⑥ 關鍵驅動 ──
    rows = [("基準", "base"),
            ("1 API 任務數成長 2027–30 取低", "api_task_lo"), ("1 API 任務數成長 2027–30 取高", "api_task_hi"),
            ("1 API 每任務 token 成長 取低", "tok_lo"), ("1 API 每任務 token 成長 取高", "tok_hi"),
            ("2 價格彈性（任務數）−1.05", "elast_strong"), ("2 價格彈性（任務數）−0.35", "elast_weak"), ("2 API 牌價年降 35%（觸發容量上限）", "price_cut"),
            ("3 每 GW 年合約價 8", "price_lo"), ("3 每 GW 年合約價 20", "price_hi"), ("3 爬坡 1 年", "ramp_lo"), ("3 爬坡 3 年", "ramp_hi"),
            ("3 AMD 合約只 3 年", "amd_short"), ("3 NonNV 產出比 TK 低", "nnv_lo"), ("3 NonNV 產出比 TK 高", "nnv_hi")]
    combos = [("任務數低＋彈性弱", "combo1"), ("任務數低＋合約價 20", "combo2"), ("彈性弱＋合約價 20＋爬坡 1 年", "combo3"),
              ("任務數低＋彈性弱＋合約價 20＋爬坡 1 年", "combo4"), ("上列＋每任務 token 低＋人數成長高＋抽成 30%", "combo5")]

    def srow(lab, sc, em=False):
        return (f'<tr class="{"em" if em else ""}"><td>{lab}</td><td>{n("REV_NetCapped", Y, "f1", sc)}</td><td>{n("COST_PropGap_VR", 2027, "sf1", sc)}</td>'
                f'<td>{n("COST_PropGap_VR", Y, "sf1", sc)}</td><td>{n("COST_Coverage", Y, "pct0", sc)}</td><td>{n("FND_ExtNeedCum", Y, "f1", sc)}</td>'
                f'<td>{n("FND_FirstGapYear", None, "yt", sc)}</td><td>{n("FND_CashNoExtTrough", None, "f1", sc)}</td></tr>')
    head = "<tr><th>情境</th><th>2030 營收淨額</th><th>2027 差額</th><th>2030 差額</th><th>2030 覆蓋率</th><th>累計外部資金</th><th>首次缺口年</th><th>現金谷底</th></tr>"
    _b = V.get("FND_ExtNeedCum", Y)
    flip = [lab for lab, sc in rows[1:] if V.get("FND_ExtNeedCum", Y, sc) > _b + 1e-9]
    flip_txt = ("<b>單一驅動落在區間端點會使累計外部資金需求高於基準者：" + "、".join(flip) + "</b>（v0.1.6、v0.2.0 時沒有任何單一驅動會產生外部資金需求）。") if flip else "<b>任何單一驅動在區間內都不會產生外部資金需求。</b>"
    S.append(f"""<section id="s6"><h2><span class="no">⑥</span>三個關鍵驅動、敏感度與「要成立必須…」</h2>
<p class="note">1＝API 需求成長（任務數、每任務 token；營收面最大驅動）；2＝價格反應（彈性、牌價年降；營收面＋容量上限）；3＝算力合約條款與產能（合約價、爬坡、AMD 期間、TPU／Trainium／AMD 產出比；成本與分母）。設定取自 Inputs 低／高欄與 Tokenomics 快照區間；全部由 Excel 重算。{flip_txt}差額單位 $B/GW/年（每 VR 等值 GW）。</p>
<div class="tw"><table><thead>{head}</thead><tbody>
{"".join(srow(lab, sc, sc == "base") for lab, sc in rows)}
</tbody></table></div>
<h3>能翻轉結論的組合</h3>
<div class="tw"><table><thead>{head}</thead><tbody>
{"".join(srow(lab, sc) for lab, sc in combos)}
</tbody></table></div>
<div class="must"><b>「按目前已到位融資與基準支出路徑，不出現最低現金缺口」要成立，必須：</b>
<ol><li><b>2029 年前不再有額外支出、2029 年後新增算力的支出不超過餘裕</b>（最大的前提；模型只計已簽合約與正式 S-1 承諾，未建「新簽合約」機制）。現金餘裕（年底現金 − 最低現金）2028 年 {n("FND_Headroom", 2028, "f1")}、2029 年只剩 {n("FND_Headroom", 2029, "f1")}、2030 年 {n("FND_Headroom", Y, "f1")}；2029 年（含）以前任何超過該年餘裕的額外支出都會產生外部資金需求。2030 年新增算力支出超過 {n("COST_GapCash", Y, "f1")}（＝當年差額）差額就轉負。對照：供給 GW 2028 → 2030 由 {n("CMP_SupplyGW", 2028, "f1")} 降到 {n("CMP_SupplyGW", Y, "f1")}，營收淨額卻由 {n("REV_NetCapped", 2028, "f1")} 增到 {n("REV_NetCapped", Y, "f1")}。</li>
<li><b>API 需求成長、價格彈性與 S-1/A 承諾的付款時點都不落在下緣</b>：v0.2.0 起 IPO 淨額進基準後，單獨落在下緣不再需要外部資金（上表），同時落在下緣才需要；API 任務數與價格彈性同時取下緣 → 2030 差額 {n("COST_PropGap_VR", Y, "sf1", "combo1")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "combo1")}；再加合約價 20、爬坡 1 年 → {n("FND_ExtNeedCum", Y, "f1", "combo4")}（首次缺口 {n("FND_FirstGapYear", None, "yt", "combo4")}）。SpaceXAI 照 S-1 路徑付滿上限（R4）→ {n("FND_ExtNeedCum", Y, "f1", "r4_on")}；未歸屬餘額回到平均攤提且快付（2027–2031）→ {n("FND_ExtNeedCum", Y, "f1", "s1u_fast")}；Google 額外 TPU 計入 → {n("FND_ExtNeedCum", Y, "f1", "gaddl_on")}。</li>
<li><b>2026 年的 ${n("FND_Committed", 2026, "f1")}B 已交割</b>（以正式 S-1 Note 4、9、10 的各輪金額計，對現金流量表特別股發行收入 {n("SRC_ANT_472", None, "f3")}，差 {n("FND_CommittedS1Gap", None, "sf2")}；S-1/A：Q3 無新股權）。v0.1.8 起最低現金改 12 個月也不需要外部資金（{n("FND_ExtNeedCum", Y, "f1", "mc_hi")}），但與 R4 同時 → {n("FND_ExtNeedCum", Y, "f1", "r4_mc_hi")}。</li></ol></div>
{V.src("REV_NetCapped", "COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_FirstGapYear", "FND_CashNoExtTrough", "COST_GapCash", "FND_Headroom", "CMP_SupplyGW")}
</section>""")

    # ── ⑦ 與 OpenAI v0.6 並排 ──
    pair = [("每 VR 等值 GW 營收淨額", "COST_PropRev_VR", "OAI_COST_PropRev_VR", "f1"), ("每 VR 等值 GW 全成本", "COST_PropFull_VR", "OAI_COST_PropFull_VR", "f1"),
            ("每 VR 等值 GW 差額（命題 1）", "COST_PropGap_VR", "OAI_COST_PropGap_VR", "sf1"), ("覆蓋率", "COST_Coverage", "OAI_COST_Coverage", "pct0"),
            ("累計外部資金需求（命題 2）", "FND_ExtNeedCum", "OAI_FND_ExtNeedCum", "f1"), ("營收淨額（截頂後）", "REV_NetCapped", "OAI_REV_NetCapped", "f1"),
            ("算力成本（現金）", "COST_Compute", "OAI_COST_Compute", "f1"), ("供給 VR 等值 GW", "CMP_Supply_VReq", "OAI_CMP_Supply_VReq", "f2"),
            ("年底現金", "FND_CashEnd", "OAI_FND_CashEnd", "f1")]
    pr = []
    for lab, a, o_, f in pair:
        em = "em" if "命題" in lab else ""
        pr.append(f'<tr class="{em}"><td>{lab}</td>' + "".join(f"<td>{cell(V, a, y, f)}<span class=\"note\">／</span>{cell(V, o_, y, f)}</td>" for y in YEARS[1:]) + "</tr>")
    S.append(f"""<section id="s7"><h2><span class="no">⑦</span>與 OpenAI v0.6.4 並排（同一合約價；Tokenomics：Anthropic {html.escape(str(tk_ver))}、OpenAI v21.7）</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>Anthropic v{ver}</span><span><i style="background:var(--oai)"></i>OpenAI v0.6.4</span><span>每 VR 等值 GW 差額，$B/GW/年；2025 兩家都是分母效應，略</span></div>
<div class="cw">{chart_pair(V)}</div>
<div class="tw"><table><thead><tr><th>Anthropic／OpenAI</th>{"".join(f"<th>{y}</th>" for y in YEARS[1:])}</tr></thead><tbody>
{"".join(pr)}
</tbody></table></div>
<p class="note"><b>每 GW 營收（v0.1.7 更新）</b>：2026 年每 VR 等值 GW 營收淨額 Anthropic {n("COST_PropRev_VR", 2026, "f1")}（v0.1.6 為 {n("COST_PropRev_VR", 2026, "f1", "v016")}；下降主要因 TK v21.7 的 TPU／Trainium 持有比下降，同額付款換得較多 GW）對 OpenAI v0.6.4 {n("OAI_COST_PropRev_VR", 2026, "f1")}，Anthropic 約為 {n("Checks!" + chk_of("f1_oai_rev26"), None, "f2")} 倍。v0.1.7 起並排快照改 OpenAI v0.6.4（其 Tokenomics 為 v21.7）；v0.1.9 起本模型改 {html.escape(str(tk_ver))}，TPU／Trainium 產出比較低使 Anthropic 的 VR 等值 GW 變少、每 VR 等值 GW 營收變高（v0.1.8：{n("COST_PropRev_VR", 2026, "f1", "v018")}），並排的這部分差異來自 Tokenomics 版本，待 OpenAI 同步 v21.10。</p>
<p class="note">差距約一半來自每 GW 營收（Anthropic 以企業與 API 為主、免費用戶負擔小，2030 年每 VR 等值 GW 營收 {n("COST_PropRev_VR", Y)} 對 {n("OAI_COST_PropRev_VR", Y)}），另一半來自算力承諾規模（2030 算力成本 {n("COST_Compute", Y)} 對 {n("OAI_COST_Compute", Y)}；Anthropic 已簽合約與 S-1 承諾 2029 年起遞減）。若 Anthropic 2029 年起比照 OpenAI 持續擴張算力，正差額會縮小或消失。兩家共用合約價 12，對 VR 世代偏低（Anthropic 雲端毛利率 2030 年 {n("COST_CloudGMPct", Y, "pct0")}）。OpenAI 數字取自 OAI_Link 快照（<code>{html.escape(str(oai_file))}</code>，SHA-256 <code>{html.escape(str(oai_sha))[:8]}…</code>），只作並排，不被任何計算頁引用。</p>
{V.src("COST_PropRev_VR", "OAI_COST_PropRev_VR", "OAI_COST_PropFull_VR", "OAI_COST_PropGap_VR", "OAI_COST_Coverage", "OAI_FND_ExtNeedCum", "OAI_REV_NetCapped", "OAI_COST_Compute", "OAI_CMP_Supply_VReq", "OAI_FND_CashEnd", "COST_CloudGMPct")}
</section>""")

    # ── ⑧ Andy 最該審的 5 項預設 ──
    S.append(f"""<section id="s8"><h2><span class="no">⑧</span>最該審的 5 項預設（取自 v{ver} 完成報告 ☆）</h2>
<ol>
<li><b>S-1/A 未歸屬餘額 {n("COST_S1Unalloc", None, "f0")} 的付款期（v0.1.8 起依 S-1/A 付款表結構；未起租 451.2 與 9-30 後 100.2 的起租年 2027、10 年為 Assumed）與只計已簽合約</b>：回到 2027–2033 平均 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "wf8_ab")}；平均且快付（2027–2031）→ 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "s1u_fast")}；起租年 2026 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "lease_lo")}。附條件或可終止 213.8 基準照計（扣除 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "r5d_on")}）。只計已簽合約使 2030 轉正（供給 GW 2030 年 {n("CMP_SupplyGW", Y, "f1")} 低於 2028 年），新增合約機制需 Andy 決定。</li>
<li><b>SpaceXAI 讀法（N1）</b>：基準維持月費 1.25 × 36＝45；S-1 路徑（爬升後每月 3.1、上限 120.5，R4）→ 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "r4_on")}（首次 {n("FND_FirstGapYear", None, "yt", "r4_on")}）、谷底 {n("FND_CashNoExtTrough", None, "f1", "r4_on")}——本版唯一單獨就能翻轉命題 2 的 S-1/A 讀法。</li>
<li><b>API 任務數成長（A2-13；與 OpenAI v0.6 同值）</b>：最大不確定；取低／高時 2030 差額 {n("COST_PropGap_VR", Y, "sf1", "api_task_lo")}／{n("COST_PropGap_VR", Y, "sf1", "api_task_hi")}、現金谷底 {n("FND_CashNoExtTrough", None, "f1", "api_task_lo")}／{n("FND_CashNoExtTrough", None, "f1", "api_task_hi")}。</li>
<li><b>2026 非算力成本與股權報酬（W4；v0.1.8 起下半年＝S-1/A Q3 實際，Q4＝Q3）</b>：非算力（不含股權報酬）{n("COST_S1fNonComp2026", None, "f1")}；股權報酬代理 {n("COST_S1fSBCProxy2026", None, "f2")}（含 Q3 非現金捐贈 1.960，Q4 亦計一次——保守；捐贈不進每人股權報酬基準，只影響命題 1 的 2026 全成本；只計一次 → 2026 差額 {n("COST_PropGap_VR", 2026, "sf1", "chr_once")}，2027–2030 不變）；整個 W4 關閉 → 2026 差額 {n("COST_PropGap_VR", 2026, "sf1", "only_w4_off")}、谷底 {n("FND_CashNoExtTrough", None, "f1", "only_w4_off")}。</li>
<li><b>最低現金月數（D17；Andy 未決）與每 GW 年合約價 12（D9）</b>：最低現金 3／6／12 個月在 v0.1.8 基準下皆不需外部資金（谷底 {n("FND_CashNoExtTrough", None, "f1")} 對 12 個月的 2029 最低現金 {n("FND_MinCash", 2029, "f1", "mc_hi")}）；合約價 20 → 谷底 {n("FND_CashNoExtTrough", None, "f1", "price_hi")}、2027 差額 {n("COST_PropGap_VR", 2027, "sf1", "price_hi")}。</li>
</ol>
<p class="note">A1–A4 共 75 項建置預設與規格 D1–D23 的全表：<code>docs/reports/20261008_v0.1.md</code> 第 ④ 節、<code>docs/reports/20261008_v0.1_對照.xlsx</code> 第 ④ 頁；v0.1.1 新增預設（U1–U8）見 <code>docs/reports/20261021_v0.1.1.md</code>；v0.1.2（Tokenomics v18.0）的語意核對與待 Andy 事項見 <code>{report}</code>。</p>
{V.src("COST_D22Under", "COST_PropGap_VR_D22", "COST_CloudGMPct")}
</section>""")

    # ── ⑨ 資料與版本 ──
    S.append(f"""<section id="s9"><h2><span class="no">⑨</span>資料與版本</h2>
<dl>
<dt>Excel</dt><dd><code>{html.escape(model.name)}</code>（v{ver}；唯一計算引擎；本頁同資料夾附 xlsx）</dd>
<dt>Tokenomics</dt><dd>{html.escape(str(tk_ver))}（<code>{html.escape(str(tk_file))}</code>），master <code>{html.escape(str(tk_sha))}</code>，讀取日 {html.escape(str(tk_date))}</dd>
<dt>OpenAI 對照</dt><dd><code>{html.escape(str(oai_file))}</code>（v0.6.4；v0.1.6 以前為 v0.6），提交 <code>{html.escape(str(oai_commit))[:7]}</code></dd>
<dt>資料日期</dt><dd>SRC_ANT 共 {src_n} 列，最新文件日期 {src_date}；模型日期 {date_disp}</dd>
<dt>檢查</dt><dd>CHK_Errors＝{n("CHK_Errors", None, "f0")}、CHK_Warnings＝{n("CHK_Warnings", None, "f0")}（C69／C70 2025／2026 年底 GW 對報導偏高——因 TK v21.7 起 TPU／Trainium 持有比下降；C89 逐家合約對帳：AMD；皆為預期的警示；v0.1.9 起 η（2025）{n("CMP_Eta2025", None, "f3")} 回到合理範圍，C54 警示消失；v0.1.6 的 C116 因自建資本支出改 GW 法而消失）；Checks {chk_ids[0]}–{chk_ids[-1]}</dd>
<dt>期間與單位</dt><dd>FY2025–FY2030 曆年制，2025 為實際校準年；$B；GW＝IT 關鍵電力；VR 等值＝以 VR200 Sol 層級每 GW 產能換算</dd>
</dl>
<div class="must"><b>來源提醒</b>：正式 S-1 已於 2026-10-26 公開（SRC_ANT_410–464；年度與上半年數字直接讀自 S-1 標 <span class="tag">Verified</span>，期中未經查核；季度、營收成本、承諾等經二手轉載標 <span class="tag">Interested-party</span>）。v0.1.7 起正式 S-1 套用於基準：2025 校準（營收、算力與基礎設施費用）、2026 上半年營收、逐家承諾與承諾三段、季度營業損益（非算力成本）、股權報酬代理（調整後 − GAAP 營業損益）。例外：AMD 維持模型 GW 型（另列 AMD＝20 情境）、6-30 現金只對照（S-1 131.65 比模型近似多 {n("FND_Cash0630Gap", None, "f1")}）。v0.1.7.1：季度數、營收成本、承諾三段、逐家金額、資本支出、集中度與 non-GAAP 調節等原為二手轉載的 46 列，經 chat 端 2026-11-02 以瀏覽器核對 SEC 原文全部一致，改 <span class="tag">Verified</span>（SRC_ANT_421 為轉載者捨入，維持 <span class="tag">Interested-party</span>、只對照）；新增 SRC_ANT_472–487（各輪融資、條件式額度、信用額度條款、受限現金）；v0.1.8：新增 SRC_ANT_488–521（S-1/A 第 1 號修訂，2026-11-02 申報：2026Q3 季度損益、承諾總額與逐年付款表、9-30 現金；chat 端 2026-11-03 以瀏覽器直讀原文，<span class="tag">Verified</span>）；SRC_ANT_384、385、445–447、450 標「已被取代」，值不改。草擬招股書的 Reuters、WSJ 轉述列（SRC_ANT 中 {src_pros} 列的來源或備註引用招股書／說明書）保留作對照，值不改。</div>
<p class="note"><span class="tag">Verified</span>已查核原文 <span class="tag">Interested-party</span>公司或利害關係方說法 <span class="tag">Analogy</span>類比推估（附區間） <span class="tag">Assumed</span>假設（附區間） <span class="tag">Derived</span>由其他數字推得 <span class="tag">Decision</span>建模決定。公司原始數據在 SRC_ANT、算力物理取自 Tokenomics（TK_Link）、假設在 Inputs、OpenAI 對照在 OAI_Link。</p>
<p class="note">本頁所有數字皆讀自 LibreOffice 重算後的 Excel（具名範圍或頁＋列 ID），HTML 不含計算；敏感度情境定義見 <code>tools/html_scenarios.yaml</code>；完整報告 <code>{report}</code>。</p>
</section>""")

    # ── 頁尾：各版更新紀錄（①之二～①之四，原在 ① 與 ② 之間） ──
    S.append('<p class="note" id="hist">以下為各版更新紀錄（①之九＝v0.2.0 在最前、①之八＝v0.1.9、①之七＝v0.1.8、①之六＝v0.1.7.1、①之五＝v0.1.7；①之二～①之四原接在 ① 之後，2026-10-31 起移到頁尾，內容不變，數字為各版當時的輸出）。</p>')
    S.extend(_hist)

    body = "\n".join(S)
    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>Anthropic 收支模型 v{ver}</title>
<style>{CSS}</style>
</head>
<body>
<main>
{body}
<footer>Anthropic 收支模型 v{ver}｜產生器 tools/build_html.py｜{date_disp}</footer>
</main>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outdir", type=Path, default=REPO / "dist")
    a = ap.parse_args()
    h, x, V = build(a.outdir)
    print(f"HTML：{h}（{h.stat().st_size:,} bytes；數字 {len(V.used)} 個）")
    print(f"Excel：{x}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""HTML 一頁摘要產生器（自 OpenAI v0.6 S6 複製；Z5 改寫內容，Z2 只改名稱）：LibreOffice 重算 → 以具名範圍或「頁!列 ID」取值 → 單一靜態 HTML。

- 計算一律由 Excel 執行：基準與敏感度情境都是「改寫 Inputs → LibreOffice 重算」後讀值；
  本程式只做取值、格式化與排版（含 SVG 幾何），不做任何模型計算。
- 敏感度情境定義在 tools/html_scenarios.yaml（值取自 Inputs 的低／高欄或倍數列）。
- 每個數字以 <span class="n" data-ref data-y data-sc data-fmt data-v> 標示來源，title 顯示 Excel 位置；
  tests/test_html.py 以 engine（pycel）逐一比對。
- 輸出單一檔案、離線可開：CSS 內嵌、圖表為內嵌 SVG、系統字型，無任何外部 URL。

用法：python3 tools/build_html.py [--outdir dist]
"""
from __future__ import annotations

import argparse
import html
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


# ── 取位：具名範圍或「頁!列 ID」→（頁, 格, 列 ID）────────────────────────────
class Locator:
    """以公式模式的活頁簿建立位址對照（不讀值）。"""

    def __init__(self, path: Path):
        wb = openpyxl.load_workbook(path)
        self.names = {k: v.attr_text for k, v in wb.defined_names.items()}
        self.rowid: dict[tuple[str, str], int] = {}       # (頁, 列 ID) → 列號
        self.rowid_of: dict[tuple[str, int], str] = {}    # (頁, 列號) → 列 ID
        for ws in wb:
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if isinstance(v, str) and re.fullmatch(r"[A-Z]{1,3}\d{1,3}|C\d{2,3}|INP_\d{3}|SRC_ZP_\d{3}", v):
                    self.rowid.setdefault((ws.title, v), r)
                    self.rowid_of[(ws.title, r)] = v
        self.inputs = wb["Inputs"]

    def locate(self, ref: str, year: int | None = None) -> tuple[str, str, str]:
        if "!" in ref:
            sheet, rid = ref.split("!", 1)
            if re.fullmatch(r"[A-Z]+\d+", rid) and (sheet, rid) not in self.rowid:   # 直接格位（README!A1 等）
                return sheet, rid, rid
            r = self.rowid[(sheet, rid)]
            col = YEAR_COL[year] if year is not None else "D"
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

    # Inputs 的值／低／高欄（常數格）
    def input_cell(self, name: str, which: str):
        sheet, coord = parse_ref(self.names[name])
        coord = coord.replace("$", "")
        r = int(re.sub(r"[A-Z]+", "", coord))
        col = {"val": "E", "lo": "F", "hi": "G"}[which]
        return self.inputs[f"{col}{r}"].value


def load_scenarios(path: Path = SCEN_FILE) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def resolve_overrides(scn_id: str, loc: Locator, cfg: dict | None = None) -> dict[str, float]:
    """情境 → {INP_nnn: 值}；值一律取自 Excel 的 Inputs（低／高欄、倍數列）。"""
    cfg = cfg or load_scenarios()
    out: dict[str, float] = {}
    for g in cfg["scenarios"][scn_id]:
        for name, spec in cfg["groups"][g]["set"].items():
            if spec in ("lo", "hi"):
                v = loc.input_cell(name, spec)
            elif isinstance(spec, dict) and "mul" in spec:
                v = loc.input_cell(name, "val") * loc.input_cell(spec["mul"], "val")
            elif isinstance(spec, (int, float)):
                v = spec
            else:
                raise ValueError(f"{g}.{name}: 不支援的寫法 {spec!r}")
            if v is None:
                raise ValueError(f"{g}.{name}: Inputs 欄位為空")
            out[name] = v
    return out


# ── 格式化（測試共用）───────────────────────────────────────────────────
def fmt(v, kind: str) -> str:
    if kind == "yr":
        return str(int(round(v)))
    if kind == "pct0":
        x = round(v * 100)
        return f"{MINUS if x < 0 else ''}{abs(x):d}%"
    if kind == "pct1":
        x = round(v * 100, 1)
        return f"{MINUS if x < 0 else ''}{abs(x):.1f}%"
    m = re.fullmatch(r"f(\d)", kind)
    if m:
        n = int(m.group(1))
        x = round(v, n)
        if x == 0:
            x = 0.0
        return f"{MINUS if x < 0 else ''}{abs(x):,.{n}f}"
    if kind == "text":
        return str(v)
    raise ValueError(kind)


# ── 值來源（LibreOffice 重算後的 xlsx）──────────────────────────────────
class Values:
    def __init__(self, loc: Locator, books: dict[str, Path]):
        self.loc = loc
        self.wb = {k: openpyxl.load_workbook(p, data_only=True) for k, p in books.items()}
        self.used: list[tuple] = []

    def get(self, ref: str, year: int | None = None, sc: str = "base"):
        sheet, coord, _ = self.loc.locate(ref, year)
        v = self.wb[sc][sheet][coord].value
        if v is None:
            raise ValueError(f"{sc}:{ref}@{year} 為空（{sheet}!{coord}）")
        if isinstance(v, str) and v.startswith("#"):
            raise ValueError(f"{sc}:{ref}@{year} 錯誤值 {v}")
        return v

    def first_year(self, ref: str, sc: str = "base") -> int:
        """該列（D–I）第一個非空值的年度（例：Funding F24 缺口年度標示）。"""
        for y in YEARS:
            sheet, coord, _ = self.loc.locate(ref, y)
            v = self.wb[sc][sheet][coord].value
            if v not in (None, ""):
                return y
        raise ValueError(f"{ref} 整列為空")

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

    def n(self, ref: str, year: int | None = None, f: str = "f1", sc: str = "base", cls: str = "") -> str:
        """可追溯的數字 span。"""
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        c = f"n {cls}".strip()
        return (f'<span class="{c}" data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" '
                f'data-v="{v!r}" title="{html.escape(self.title(ref, year, sc))}">{fmt(v, f)}</span>')

    def t(self, ref: str, year: int | None = None, f: str = "f1", sc: str = "base", **attrs) -> str:
        """SVG <text>，同樣帶 data-ref（測試一併比對）。"""
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        a = " ".join(f'{k.replace("_", "-")}="{val}"' for k, val in attrs.items())
        return (f'<text {a} data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" data-v="{v!r}">'
                f'<title>{html.escape(self.title(ref, year, sc))}</title>{fmt(v, f)}</text>')

    def src(self, *refs: str) -> str:
        """圖表下方小字：Excel 位置（頁＋列 ID，具名範圍附註）。"""
        out = []
        for r in refs:
            multi = "!" not in r and ":" in self.loc.names.get(r, "")
            sheet, _, rid = self.loc.locate(r, YEARS[0] if ("!" in r or multi) else None)
            out.append(f"{sheet} {rid}" + ("" if "!" in r else f"（{r}）"))
        return '<p class="src">Excel：' + "、".join(out) + "</p>"


# ── SVG 小工具 ─────────────────────────────────────────────────────────
def nice_max(x: float) -> float:
    """座標軸上限：取 {1, 1.2, 1.6, 2, 2.4, 4, 6, 8, 10} × 10^k 中第一個 ≥ x 的值（4 等分刻度皆為整齊值）。"""
    p = 10 ** math.floor(math.log10(x))
    for m in (1, 1.2, 1.6, 2, 2.4, 4, 6, 8, 10):
        if m * p >= x - 1e-12:
            return m * p
    return 10 * p


def axis(W, H, L, R, T, B, ymin, ymax, ticks, unit=""):
    out = []
    for i in range(ticks + 1):
        val = ymin + (ymax - ymin) * i / ticks
        y = T + (H - T - B) * (1 - (val - ymin) / (ymax - ymin))
        out.append(f'<line class="grid{" zero" if abs(val) < 1e-12 else ""}" x1="{L}" x2="{W - R}" y1="{y:.1f}" y2="{y:.1f}"/>')
        lab = f"{MINUS if val < 0 else ''}{abs(val):g}"
        out.append(f'<text class="tick" x="{L - 6}" y="{y + 4:.1f}" text-anchor="end">{lab}</text>')
    if unit:
        out.append(f'<text class="tick" x="2" y="{T - 10}">{unit}</text>')
    return out


def ymap(val, H, T, B, ymin, ymax):
    return T + (H - T - B) * (1 - (val - ymin) / (ymax - ymin))


# ── 頁面 ───────────────────────────────────────────────────────────────
CSS = r"""
:root{--bg:#fbfaf7;--fg:#1d2329;--muted:#5d6670;--line:#d9dde1;--card:#ffffff;--accent:#0d5c63;
--rev:#2563a8;--cost:#c2410c;--gap:#b42318;--sub:#2563a8;--api:#5aa0d8;--ads:#a3c9ea;--net:#1d2329;--tgt:#7a3e9d;
--inf:#2f855a;--rd:#9ae6b4;--con:#4a5568;--own:#a0aec0;--amz:#d69e2e;--warn-bg:#fff4e5;--ok:#2f855a}
@media (prefers-color-scheme: dark){:root{--bg:#14181c;--fg:#e8ebee;--muted:#9aa4ae;--line:#323a42;--card:#1b2127;--accent:#5fc4cb;
--rev:#6aa7ea;--cost:#f08a4b;--gap:#ff7b72;--sub:#6aa7ea;--api:#3f7fbf;--ads:#9cc7ee;--net:#e8ebee;--tgt:#c792ea;
--inf:#56c288;--rd:#2e6b4a;--con:#a0aec0;--own:#5a6573;--amz:#e6b450;--warn-bg:#2b2416;--ok:#56c288}}
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
svg{width:100%;height:auto;display:block}
svg text{fill:var(--fg);font-size:12px;font-family:inherit}
svg .tick{fill:var(--muted);font-size:11px}
svg .grid{stroke:var(--line);stroke-width:1}
svg .grid.zero{stroke:var(--muted)}
svg .lbl{font-size:11.5px}
svg .gaplbl{fill:var(--gap);font-weight:700;font-size:12px}
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
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    rev = [V.get("COST_PropRev_VR", y) for y in YEARS]
    full = [V.get("COST_PropFull_VR", y) for y in YEARS]
    ymax = nice_max(max(full + rev) * 1.08)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="每 VR 等值 GW 營收與全成本 2025–2030">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B/GW/年")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.3
    for i, y in enumerate(YEARS):
        x0 = L + slot * i + slot * 0.18
        for j, (vals, cls) in enumerate(((rev, "rev"), (full, "cost"))):
            v = vals[i]
            yt = ymap(v, H, T, B, 0, ymax)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
        yt = ymap(full[i], H, T, B, 0, ymax)
        o.append(V.t("COST_PropGap_VR", y, "f1", x=f"{x0 + bw:.1f}", y=f"{yt - 6:.1f}", text_anchor="middle", **{"class": "gaplbl"}))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def chart_rev(V: Values) -> str:
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    tgt = [V.get("RVS_Target", y) for y in YEARS]
    ymax = nice_max(max(tgt) * 1.05)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="營收結構與管理層目標">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.36
    for i, y in enumerate(YEARS):
        xc = L + slot * i + slot * 0.42
        base = 0.0
        for ref, cls in (("Reverse!X06", "sub"), ("Reverse!X07", "api"), ("Reverse!X08", "ads"), ("Reverse!X09", "ads")):
            v = V.get(ref, y)
            if v <= 0:
                continue
            y1, y0 = ymap(base + v, H, T, B, 0, ymax), ymap(base, H, T, B, 0, ymax)
            o.append(f'<rect x="{xc - bw / 2:.1f}" y="{y1:.1f}" width="{bw:.1f}" height="{y0 - y1:.1f}" style="fill:var(--{cls})"/>')
            base += v
        net = V.get("REV_NetCapped", y)
        yn = ymap(net, H, T, B, 0, ymax)
        o.append(f'<line x1="{xc - bw / 2 - 4:.1f}" x2="{xc + bw / 2 + 4:.1f}" y1="{yn:.1f}" y2="{yn:.1f}" style="stroke:var(--net);stroke-width:2.5"/>')
        o.append(V.t("REV_GrossCapped", y, "f1", x=f"{xc:.1f}", y=f"{ymap(base, H, T, B, 0, ymax) - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        if y >= 2026:
            yt = ymap(tgt[i], H, T, B, 0, ymax)
            xd = xc + bw / 2 + 8
            o.append(f'<path d="M{xd:.1f},{yt - 6:.1f} l6,6 l-6,6 l-6,-6 z" style="fill:var(--tgt)"/>')
            o.append(V.t("RVS_Target", y, "f0", x=f"{xd + 9:.1f}", y=f"{yt + 4:.1f}", **{"class": "lbl"}, style="fill:var(--tgt)"))
        o.append(f'<text class="tick" x="{xc:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
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
    W, H, L, R, T, B = 640, 280, 44, 10, 26, 34
    yrs = YEARS[1:]
    base = [V.get("FND_ExtNeed", y) for y in yrs]
    amz = [V.get("Funding!F31", y) for y in yrs]
    ymax = nice_max(max(base + amz) * 1.12)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="各年外部資金需求：基準與 Amazon 情境">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "$B")
    slot = (W - L - R) / len(yrs)
    bw = slot * 0.3
    for i, y in enumerate(yrs):
        x0 = L + slot * i + slot * 0.18
        for j, (ref, cls) in enumerate((("FND_ExtNeed", "gap"), ("Funding!F31", "amz"))):
            v = V.get(ref, y)
            yt = ymap(v, H, T, B, 0, ymax)
            if v > 0:
                o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
            o.append(V.t(ref, y, "f1", x=f"{x0 + j * bw + bw / 2 - 1:.1f}", y=f"{yt - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def year_table(V: Values, rows: list[tuple], years=YEARS, head="$B") -> str:
    """rows：(標籤, ref, fmt, 樣式)；ref 為 None 時整列留空標題。"""
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


def build(outdir: Path) -> tuple[Path, Path, Values]:
    from parity_lib import lo_recalc, set_inputs

    model = current_model_path()
    loc = Locator(model)
    cfg = load_scenarios()
    tmp = Path(tempfile.mkdtemp(prefix="s6_html_"))
    books = {}
    for sc in cfg["scenarios"]:
        src = tmp / f"in_{sc}" / model.name
        src.parent.mkdir(parents=True)
        ov = resolve_overrides(sc, loc, cfg)
        if ov:
            set_inputs(model, src, ov)
        else:
            shutil.copy(model, src)
        books[sc] = lo_recalc(src, tmp / f"out_{sc}")
    V = Values(loc, books)
    for sc in books:
        if V.get("CHK_Errors", None, sc) != 0 and sc == "base":
            raise SystemExit("基準 CHK_Errors ≠ 0，不產生成品")

    stamp = model.name.split("_")[0]
    ver = re.search(r"_v([\d.]+)\.xlsx$", model.name).group(1)
    html_path = outdir / f"{stamp}_智譜收支模型_v{ver.replace('.', '_')}.html"
    xlsx_path = outdir / f"{stamp}_智譜收支模型_v{ver.replace('.', '_')}.xlsx"
    page = render(V, model, ver, stamp)
    outdir.mkdir(parents=True, exist_ok=True)
    html_path.write_text(page, encoding="utf-8")
    shutil.copyfile(model, xlsx_path)
    shutil.rmtree(tmp, ignore_errors=True)
    return html_path, xlsx_path, V


def latest_src_date(model: Path) -> str:
    ws = openpyxl.load_workbook(model, read_only=True)["SRC_ZP"]
    ds = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        d = row[8]
        if isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", d.strip()):
            ds.append(d.strip())
    return max(ds) if ds else "—"


def render(V: Values, model: Path, ver: str, stamp: str) -> str:
    n = V.n
    Y = 2030
    gap_first = V.first_year("Funding!F24")
    peak = V.first_year("Funding!F27")
    date_disp = f"{stamp[:4]}-{stamp[4:6]}-{stamp[6:]}"
    tk_file = V.get("TK_Link!B3")
    tk_ver = V.get("TK_Link!B4")
    tk_sha = V.get("TK_Link!B5")
    tk_date = V.get("TK_Link!B6")

    S = []
    # ── 標頭 ──
    S.append(f"""<header><h1>智譜收支模型 v{ver}｜一頁摘要</h1>
<div class="meta">{date_disp}｜Excel <code>{html.escape(model.name)}</code>｜Tokenomics {html.escape(str(tk_ver))}（<code>{html.escape(str(tk_sha))[:7]}</code>）｜單位 $B（十億美元），曆年制｜數字游標停留可見 Excel 位置；手機上圖表可左右滑動</div></header>""")

    # ── ① 命題與結論 ──
    S.append(f"""<section id="s1"><h2><span class="no">①</span>命題與一句話結論</h2>
<p class="note">命題：智譜每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本？若不能，缺口要多少外部資金、由誰補？</p>
<p class="lead">不能。基準下 2025–2030 每一年每 VR 等值 GW 營收都低於全成本（2030 年 {n("COST_PropRev_VR", Y)} 對 {n("COST_PropFull_VR", Y)} $B/GW/年，覆蓋率 {n("COST_Coverage", Y, "pct0")}）；
2026-03 輪已到位的 ${n("FND_Committed", 2026, "f0")}B 撐不過 {n("Funding!F24", gap_first, "yr")} 年，2027–2030 累計需外部資金 ${n("FND_ExtNeedCum", Y, "f0")}B（峰值 {n("Funding!F27", peak, "yr")} 年 ${n("FND_ExtNeed", peak, "f0")}B）。</p>
<div class="kpis">
<div class="kpi"><div class="k">2030 每 VR 等值 GW：營收 vs 全成本</div><div class="v">{n("COST_PropRev_VR", Y)} <span class="note">vs</span> {n("COST_PropFull_VR", Y)}</div><div class="s">$B/GW/年；差額 <b class="neg">{n("COST_PropGap_VR", Y)}</b></div></div>
<div class="kpi"><div class="k">2030 覆蓋率（營收 ÷ 全成本）</div><div class="v">{n("COST_Coverage", Y, "pct0")}</div><div class="s">2025 為 {n("COST_Coverage", 2025, "pct0")}；全成本含股權報酬、現金口徑</div></div>
<div class="kpi"><div class="k">累計外部資金需求 2027–2030</div><div class="v neg">{n("FND_ExtNeedCum", Y, "f0")}</div><div class="s">$B；峰值 {n("Funding!F27", peak, "yr")} 年 {n("FND_ExtNeed", peak, "f1")}</div></div>
<div class="kpi"><div class="k">首次需要外部資金的年度</div><div class="v">{n("Funding!F24", gap_first, "yr")}</div><div class="s">已到位 {n("FND_Committed", 2026, "f0")}（2026 到位）只撐到前一年底</div></div>
</div>
<p class="note">即使管理層營收目標（2030 年 ${n("RVS_Target", Y, "f0")}B）全數達成、支出不變，仍需外部資金 ${n("RVS_ExtNeedCum", Y, "f1")}B（反向模式，只作對照）。Amazon 條件式 $35B 若到位，累計需求降為 {n("FND_ExtNeedCumAmzn", Y, "f1")}。</p>
</section>""")

    # ── ② 每 VR 等值 GW ──
    S.append(f"""<section id="s2"><h2><span class="no">②</span>每 VR 等值 GW：營收 vs 全成本（2025–2030）</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>營收（截頂後淨額）</span><span><i style="background:var(--cost)"></i>全成本（算力＋非算力＋股權報酬，現金口徑）</span><span><b style="color:var(--gap)">紅字</b>＝差額</span></div>
<div class="cw">{chart_gw(V)}</div>
{year_table(V, [("營收", "COST_PropRev_VR", "f1", ""), ("　算力成本（合約實付＋自有資本支出）", "COST_PropCompute_VR", "f1", ""),
                ("　非算力成本（不含股權報酬）", "Cost!K79", "f1", ""), ("　股權報酬", "Cost!K80", "f1", ""),
                ("全成本", "COST_PropFull_VR", "f1", ""), ("差額", "COST_PropGap_VR", "f1", "em neg"),
                ("覆蓋率", "COST_Coverage", "pct0", ""), ("分母：供給 VR 等值 GW", "CMP_Supply_VReq", "f2", "")], head="$B/GW/年")}
<p class="note">差額逐年收斂（算力單價隨世代下降、分母擴大），但到 2030 仍為負；經濟口徑（自建算力以持有成本年化）2030 差額為 {n("Cost!K88", Y)}。2025 分母只有 {n("CMP_Supply_VReq", 2025, "f2")} VR 等值 GW，每 GW 數字偏大。</p>
{V.src("COST_PropRev_VR", "COST_PropFull_VR", "COST_PropGap_VR", "COST_Coverage", "COST_PropCompute_VR", "Cost!K79", "Cost!K80", "CMP_Supply_VReq", "Cost!K88")}
</section>""")

    # ── ③ 營收結構 ──
    S.append(f"""<section id="s3"><h2><span class="no">③</span>營收結構與管理層目標</h2>
<div class="legend"><span><i style="background:var(--sub)"></i>訂閱</span><span><i style="background:var(--api)"></i>API</span><span><i style="background:var(--ads)"></i>廣告＋其他</span><span><i style="background:var(--net);height:3px"></i>淨額（扣 Microsoft 分成）</span><span><i style="background:var(--tgt);transform:rotate(45deg)"></i>管理層目標</span></div>
<div class="cw">{chart_rev(V)}</div>
{year_table(V, [("訂閱", "Reverse!X06", "f1", ""), ("API", "Reverse!X07", "f1", ""), ("廣告", "Reverse!X08", "f1", ""), ("其他", "Reverse!X09", "f1", ""),
                ("總額", "REV_GrossCapped", "f1", "em"), ("Microsoft 分成", "REV_MSShareCapped", "f1", ""), ("淨額", "REV_NetCapped", "f1", "em"),
                ("管理層目標（只對照）", "RVS_Target", "f1", ""), ("差距（目標 − 正向總額）", "RVS_Gap", "f1", "")])}
<p class="note">2030 正向總額 {n("REV_GrossCapped", Y)} 對管理層目標 {n("RVS_Target", Y, "f0")}，差距 {n("RVS_Gap", Y)}；2026 年模型 {n("REV_GrossCapped", 2026)} 對目標 {n("RVS_Target", 2026, "f0")}。若只靠 API 補足需 {n("RVS_MultAPI", Y, "f2")} 倍、只靠訂閱需 {n("RVS_MultSub", Y, "f2")} 倍、兩線等比例需 {n("RVS_MultProp", Y, "f2")} 倍。Microsoft 分成以總額 20% 計、累計上限 {n("REV_MSCumCapped", Y, "f0")}，觸頂後不再扣（2030 年分成 {n("REV_MSShareCapped", Y, "f0")}）。廣告為獨立慢成長一列，不進核心分析。</p>
{V.src("Reverse!X06", "Reverse!X07", "Reverse!X08", "REV_GrossCapped", "REV_MSShareCapped", "REV_NetCapped", "RVS_Target", "RVS_Gap", "RVS_MultAPI", "RVS_MultSub", "RVS_MultProp", "REV_MSCumCapped")}
</section>""")

    # ── ④ 算力 ──
    S.append(f"""<section id="s4"><h2><span class="no">④</span>算力：總需求 vs 合約＋自建供給（GW，IT 電力）</h2>
<div class="legend"><span><i style="background:var(--inf)"></i>有效推論 GW</span><span><i style="background:var(--rd)"></i>研發 GW（殘差）</span><span><i style="background:var(--con)"></i>合約供給</span><span><i style="background:var(--own)"></i>自建供給</span><span>每年左＝需求、右＝供給</span></div>
<div class="cw">{chart_compute(V)}</div>
{year_table(V, [("有效推論 GW", "CMP_InfGW_Eff", "f2", ""), ("研發 GW（殘差）", "CMP_RDGW", "f2", ""), ("總需求 GW", "CMP_DemandGW", "f2", "em"),
                ("合約供給 GW", "CMP_SupplyContractGW", "f2", ""), ("自建供給 GW", "CMP_Supply_Owned", "f2", ""), ("供給 GW 合計", "CMP_SupplyGW", "f2", "em"),
                ("推論可用 GW（容量上限）", "CMP_InfAvailGW", "f2", "")], head="GW")}
<p><b>研發 GW 是殘差，不是獨立需求。</b>模型先由 token 需求推出推論 GW，再把「供給扣閒置後剩下的」全部記為研發；2027 年起研發占總需求約 6–9 成（2027 年 {n("CMP_RDGW", 2027, "f1")} / {n("CMP_DemandGW", 2027, "f1")} GW，2030 年 {n("CMP_RDGW", Y, "f1")} / {n("CMP_DemandGW", Y, "f1")} GW）。意義：①總需求其實由已簽合約決定，不是由用量推出；②若這些 GW 實際上是閒置或轉賣，現金成本照付、命題結論不變，但「研發投入」的解讀要打折；③推論只用掉 2030 年推論可用 GW 的約一半（{n("CMP_InfGW_Eff", Y, "f1")} / {n("CMP_InfAvailGW", Y, "f1")}），營收上行的算力餘裕有限。</p>
{V.src("CMP_InfGW_Eff", "CMP_RDGW", "CMP_DemandGW", "CMP_SupplyContractGW", "CMP_Supply_Owned", "CMP_SupplyGW", "CMP_InfAvailGW")}
</section>""")

    # ── ⑤ 現金與外部資金 ──
    S.append(f"""<section id="s5"><h2><span class="no">⑤</span>現金與外部資金需求</h2>
<div class="legend"><span><i style="background:var(--gap)"></i>當年外部資金需求（基準）</span><span><i style="background:var(--amz)"></i>情境：Amazon 條件式 $35B 於 2026 到位</span></div>
<div class="cw">{chart_cash(V)}</div>
{year_table(V, [("營收淨額", "Funding!F01", "f1", ""), ("算力成本（現金）", "Funding!F02", "f1", ""), ("非算力成本（不含股權報酬）", "Funding!F03", "f1", ""),
                ("自由現金流（股權報酬加回）", "FND_FCF", "f1", "em"), ("已到位融資", "FND_Committed", "f1", ""), ("當年外部資金需求", "FND_ExtNeed", "f1", "em"),
                ("累計外部資金需求", "FND_ExtNeedCum", "f1", "em"), ("年底現金（最低現金 10）", "FND_CashEnd", "f1", ""),
                ("只靠已到位融資的年底現金", "FND_CashNoExt", "f1", ""), ("情境：Amazon 當年外部資金需求", "Funding!F31", "f1", ""),
                ("情境：Amazon 累計外部資金需求", "FND_ExtNeedCumAmzn", "f1", "")])}
<p class="note">來源順序：①期初現金（2025 年底 {n("FND_CashEnd", 2025, "f0")}）→ ②已到位股權（2026-03 輪無條件部分 {n("FND_Committed", 2026, "f0")}）→ ③外部資金（新股權、債務或延後承諾；不分種類、不計利息）。<b>或有負債（只列示、不作資金來源）</b>：Nvidia 擔保與晶片融資合計 {n("FND_Contingent", None, "f0")}；表外承諾存量 {n("Funding!F37", None, "f0")}（S-1 草案轉述）；合約 2030 年後剩餘承諾 {n("FND_CommitAfter2030", None, "f1")}。</p>
{V.src("Funding!F01", "Funding!F02", "Funding!F03", "FND_FCF", "FND_Committed", "FND_ExtNeed", "FND_ExtNeedCum", "FND_CashEnd", "FND_CashNoExt", "Funding!F31", "FND_ExtNeedCumAmzn", "FND_Contingent", "Funding!F37", "FND_CommitAfter2030")}
</section>""")

    # ── ⑥ 關鍵驅動與敏感度 ──
    rows = [("基準", "base"),
            ("1 API 任務數成長 2027–30 ×1.5", "api_task_hi"), ("1 API 任務數成長 2027–30 ×0.5", "api_task_lo"),
            ("1 價格彈性（任務）−1.05", "elast_strong"), ("1 價格彈性（任務）−0.35", "elast_weak"),
            ("2 自有資本支出 ×0.6", "capex_lo"), ("2 自有資本支出 ×1.5", "capex_hi"), ("2 自有資本支出歸零", "capex_zero"),
            ("3 員工人數成長 取低", "hc_lo"), ("3 員工人數成長 取高", "hc_hi"),
            ("並列：合約價 8（只改每 GW 指標）", "price_lo"), ("並列：合約價 20", "price_hi"),
            ("組合：有利", "favorable"), ("組合：有利＋自有資本支出歸零", "favorable_capex0"), ("組合：不利", "unfavorable")]
    tr = []
    for lab, sc in rows:
        tr.append(f'<tr class="{"em" if sc == "base" else ""}"><td>{lab}</td><td>{n("REV_NetCapped", Y, "f1", sc)}</td>'
                  f'<td>{n("COST_PropGap_VR", Y, "f1", sc)}</td><td>{n("COST_Coverage", Y, "pct0", sc)}</td><td>{n("FND_ExtNeedCum", Y, "f1", sc)}</td></tr>')
    S.append(f"""<section id="s6"><h2><span class="no">⑥</span>三個關鍵驅動與敏感度</h2>
<p class="note">1＝API 需求成長（營收面最大驅動）；2＝自有資本支出（現金面最大驅動，v0.5 Assumed 每年 20→70）；3＝員工人數成長（非算力成本與股權報酬）。設定取自 Inputs 低／高欄與情境倍數列；全部由 Excel 重算。</p>
<div class="tw"><table><thead><tr><th>情境</th><th>2030 營收淨額</th><th>2030 每 VR GW 差額</th><th>2030 覆蓋率</th><th>累計外部資金需求</th></tr></thead><tbody>
{"".join(tr)}
</tbody></table></div>
<p class="note">有利組合＝API ×1.5＋彈性 −1.05＋廣告取高＋自有資本支出 ×0.6＋人數取低＋合約價 8；不利組合＝API ×0.5＋彈性 −0.35＋自有資本支出 ×1.5＋人數取高＋合約價 20。</p>
<div class="must"><b>要讓命題成立（每 GW 營收 ≥ 全成本），必須：</b>
<ol><li>2030 年營收淨額至少達到當年全成本 {n("COST_FullCash", Y, "f1")}（基準只有 {n("REV_NetCapped", Y, "f1")}，覆蓋率 {n("COST_Coverage", Y, "pct0")}）——相當於管理層 2030 目標 {n("RVS_Target", Y, "f0")} 的大部分要實現；</li>
<li>而且新增營收要靠每 token 價格／ARPU，或同步增加推論算力：2030 推論可用 {n("CMP_InfAvailGW", Y, "f1")} GW，基準已用 {n("CMP_InfGW_Eff", Y, "f1")} GW；有利組合下用量放大即觸發容量上限（係數 {n("CMP_CapFactor", Y, "f2", "favorable")}），未截頂淨額 {n("REV_Net", Y, "f1", "favorable")} 只實得 {n("REV_NetCapped", Y, "f1", "favorable")}。</li></ol>
沒有任何單一驅動、也沒有「全部有利」的組合能讓任何一年的差額轉正；即使有利組合＋不建自有算力，仍需外部資金 {n("FND_ExtNeedCum", Y, "f1", "favorable_capex0")}。</div>
{V.src("REV_NetCapped", "COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "COST_FullCash", "CMP_InfAvailGW", "CMP_CapFactor", "REV_Net")}
</section>""")

    # ── ⑦ Andy 最該審的 5 項預設 ──
    S.append(f"""<section id="s7"><h2><span class="no">⑦</span>最該審的 5 項預設（取自 v0.6 完成報告 ☆）</h2>
<ol>
<li><b>研發 GW 是殘差（S3 #15）</b>：合約供給扣閒置後減推論全部算研發，2027–2030 研發 {n("CMP_RDGW", 2027, "f1")}／{n("CMP_RDGW", 2028, "f1")}／{n("CMP_RDGW", 2029, "f1")}／{n("CMP_RDGW", 2030, "f1")} GW。替代：以 Tokenomics <code>IF_AllocRDGW</code> 或計畫算力推出獨立研發需求。</li>
<li><b>自有資本支出路徑 20→70／年與自建 GW 計入分母（S4 #2）</b>：現金面最大單一驅動，無公開來源；歸零時累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1")} → {n("FND_ExtNeedCum", Y, "f1", "capex_zero")}。</li>
<li><b>API 需求成長與價格彈性（S2 #13：2026 成長率 134% 不重校準）</b>：2030 營收淨額區間 {n("REV_NetCapped", Y, "f1", "elast_weak")}–{n("REV_NetCapped", Y, "f1", "elast_strong")}；2026 模型總額 {n("REV_Gross", 2026, "f1")} 高於獨立推估 {n("Revenue!R41", 2026, "f1")}。</li>
<li><b>Microsoft 分成在現金流中扣除（S5 #5）、2025 算力成本用合約實付而非實際（S4 #7）</b>：分成在財報中的位置未揭露，若已含在成本內則累計需求高估至多約分成累計上限 {n("REV_MSCumCapped", Y, "f0")}；2025 合約實付 {n("COST_Compute", 2025, "f2")} 對實際 {n("COST_Actual2025", None, "f1")}。</li>
<li><b>員工人數成長路徑（S4 #16）與股權報酬自費用扣出單列（S4 #13）</b>：人數為 Assumed（2026–2030 年增 0.55／0.30／0.20／0.15／0.10）；取低／高時累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "hc_lo")}／{n("FND_ExtNeedCum", Y, "f1", "hc_hi")}。</li>
</ol>
<p class="note">S1–S5 共 99 項預設的全表：<code>docs/reports/20261008_v0.6-P5.md</code>、<code>20261008_v0.6-P5_對照.xlsx</code> 第 ④ 頁。</p>
</section>""")

    # ── ⑧ 資料與版本 ──
    S.append(f"""<section id="s8"><h2><span class="no">⑧</span>資料與版本</h2>
<dl>
<dt>Excel</dt><dd><code>{html.escape(model.name)}</code>（v{ver}；唯一計算引擎；本頁同資料夾附 xlsx）</dd>
<dt>Tokenomics</dt><dd>{html.escape(str(tk_ver))}（<code>{html.escape(str(tk_file))}</code>），master <code>{html.escape(str(tk_sha))}</code>，讀取日 {html.escape(str(tk_date))}</dd>
<dt>資料日期</dt><dd>SRC_ZP 最新文件日期 {latest_src_date(model)}；模型日期 {date_disp}</dd>
<dt>檢查</dt><dd>CHK_Errors＝{n("CHK_Errors", None, "f0")}（Checks C01–C130）</dd>
<dt>期間與單位</dt><dd>FY2025–FY2030 曆年制，2025 為實際校準年；$B；GW＝IT 關鍵電力；VR 等值＝以 VR200 Sol 層級每 GW 產能換算</dd>
</dl>
<p class="note"><span class="tag">Verified</span>已查核原文 <span class="tag">Interested-party</span>公司或利害關係方說法 <span class="tag">Analogy</span>類比推估（附區間） <span class="tag">Assumed</span>假設（附區間） <span class="tag">Derived</span>由其他數字推得 <span class="tag">Decision</span>建模決定。公司原始數據在 SRC_ZP、算力物理取自 Tokenomics（TK_Link）、假設在 Inputs。</p>
<p class="note">本頁所有數字皆讀自 LibreOffice 重算後的 Excel（具名範圍或頁＋列 ID），HTML 不含計算；敏感度情境定義見 <code>tools/html_scenarios.yaml</code>。</p>
</section>""")

    body = "\n".join(S)
    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>智譜收支模型 v{ver}</title>
<style>{CSS}</style>
</head>
<body>
<main>
{body}
<footer>智譜收支模型 v{ver}｜產生器 tools/build_html.py｜{date_disp}</footer>
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

#!/usr/bin/env python3
"""HTML 一頁摘要產生器（Z5；以 OpenAI v0.6 S6 版改寫為智譜）：LibreOffice 重算 → 以具名範圍或「頁!列 ID」取值 → 單一靜態 HTML。

- 計算一律由 Excel 執行：基準與敏感度情境都是「改寫 Inputs／SRC_ZP 輸入格 → LibreOffice 重算」後讀值；
  本程式只做取值、格式化與排版（含 SVG 幾何），不做任何模型計算。
- 敏感度情境定義在 tools/html_scenarios.yaml（值取自 Inputs 的低／高欄，翻轉點取自 v0.1 敏感度 JSON）。
- 每個數字以 <span class="n" data-ref data-y data-sc data-fmt data-v> 標示來源，title 顯示 Excel 位置；
  tests/test_html.py 以 engine（pycel）逐一比對。
- 輸出單一檔案、離線可開：CSS 內嵌、圖表為內嵌 SVG、系統字型，無任何外部 URL。
- 幣別：公司金額 RMB 億（與財報一致）；與 OpenAI 並排一節用 Excel 的美元口徑列（Cost 第七、九節、Funding 第十節、OAI_Link）。

用法：python3 tools/build_html.py [--outdir dist]
"""
from __future__ import annotations

import argparse
import html
import json
import math
import re
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
import yaml
from openpyxl.utils import column_index_from_string, get_column_letter

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests" / "parity"))
from engine import current_model_path, parse_ref  # noqa: E402

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
# 各頁共同版面：D–I＝2025–2030；K＝1H26、L＝2H26
YEAR_COL: dict = {y: chr(ord("D") + i) for i, y in enumerate(YEARS)}
YEAR_COL.update({"1H26": "K", "2H26": "L"})
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

    def locate(self, ref: str, year=None) -> tuple[str, str, str]:
        if "!" in ref:
            sheet, rid = ref.split("!", 1)
            if re.fullmatch(r"[A-Z]+\d+", rid) and (sheet, rid) not in self.rowid:   # 直接格位（TK_Link!B3 等）
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
            if isinstance(year, int):        # 依範圍起始欄位移（OAI_Link 為 F–K）
                c0 = column_index_from_string(re.sub(r"\d+", "", a))
                coord = f"{get_column_letter(c0 + YEARS.index(year))}{r}"
            else:
                coord = f"{YEAR_COL[year]}{r}"
        r = int(re.sub(r"[A-Z]+", "", coord))
        return sheet, coord, self.rowid_of.get((sheet, r), "")

    # Inputs 的值／低／高欄（常數格；智譜 Inputs 版面：F＝值、G＝低、H＝高）
    def input_cell(self, name: str, which: str):
        sheet, coord = parse_ref(self.names[name])
        coord = coord.replace("$", "")
        r = int(re.sub(r"[A-Z]+", "", coord))
        col = {"val": "F", "lo": "G", "hi": "H"}[which]
        return self.inputs[f"{col}{r}"].value


def load_scenarios(path: Path = SCEN_FILE) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def _flip_values(cfg: dict) -> dict[str, float]:
    d = json.loads((REPO / cfg["flips_json"]).read_text(encoding="utf-8"))
    return {f["id"]: f["value"] for f in d["flips"]}


def resolve_overrides(scn_id: str, loc: Locator, cfg: dict | None = None) -> dict[str, float]:
    """情境 → {INP_nnn／SRC_ZP_nnn: 值}；值取自 Excel 的 Inputs（低／高欄）或敏感度 JSON 的翻轉點。"""
    cfg = cfg or load_scenarios()
    out: dict[str, float] = {}
    flips = None
    for name, spec in cfg["scenarios"][scn_id]["set"].items():
        if spec in ("lo", "hi"):
            v = loc.input_cell(name, spec)
        elif isinstance(spec, dict) and "flip" in spec:
            flips = flips or _flip_values(cfg)
            v = flips[spec["flip"]]
        elif isinstance(spec, (int, float)) and spec == 0:
            v = 0
        else:
            raise ValueError(f"{scn_id}.{name}: 不支援的寫法 {spec!r}")
        if v is None:
            raise ValueError(f"{scn_id}.{name}: 值為空")
        out[name] = v
    return out


# ── 格式化（測試共用）───────────────────────────────────────────────────
def fmt(v, kind: str) -> str:
    if kind in ("yr", "yrt"):            # yrt：年度或文字（例：「期間內無」）
        if isinstance(v, str):
            return v
        return str(int(round(v)))
    if kind == "text":
        return str(v)
    if kind == "hm2y":                    # 單位顯示：百萬 → 億（÷100），一位小數
        return fmt(v / 100, "f1")
    if kind == "pct0":
        x = round(v * 100)
        return f"{MINUS if x < 0 else ''}{abs(x):d}%"
    if kind == "pct2":
        x = round(v * 100, 2)
        return f"{MINUS if x < 0 else ''}{abs(x):.2f}%"
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
    raise ValueError(kind)


# ── 值來源（LibreOffice 重算後的 xlsx）──────────────────────────────────
class Values:
    def __init__(self, loc: Locator, books: dict[str, Path]):
        self.loc = loc
        self.wb = {k: openpyxl.load_workbook(p, data_only=True) for k, p in books.items()}
        self.used: list[tuple] = []

    def get(self, ref: str, year=None, sc: str = "base"):
        sheet, coord, _ = self.loc.locate(ref, year)
        v = self.wb[sc][sheet][coord].value
        if v is None:
            raise ValueError(f"{sc}:{ref}@{year} 為空（{sheet}!{coord}）")
        if isinstance(v, str) and v.startswith("#"):
            raise ValueError(f"{sc}:{ref}@{year} 錯誤值 {v}")
        return v

    def title(self, ref: str, year, sc: str) -> str:
        sheet, coord, rid = self.loc.locate(ref, year)
        parts = [f"{sheet} {rid}".strip() if rid else sheet, coord]
        if "!" not in ref and ref != rid:
            parts.append(ref)
        if year is not None:
            parts.append(str(year))
        if sc != "base":
            parts.append(f"情境 {sc}")
        return "｜".join(parts)

    def n(self, ref: str, year=None, f: str = "f1", sc: str = "base", cls: str = "") -> str:
        """可追溯的數字 span。"""
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        c = f"n {cls}".strip()
        return (f'<span class="{c}" data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" '
                f'data-v="{html.escape(repr(v))}" title="{html.escape(self.title(ref, year, sc))}">{fmt(v, f)}</span>')

    def t(self, ref: str, year=None, f: str = "f1", sc: str = "base", **attrs) -> str:
        """SVG <text>，同樣帶 data-ref（測試一併比對）。"""
        v = self.get(ref, year, sc)
        self.used.append((ref, year, sc, f))
        y = "" if year is None else f' data-y="{year}"'
        a = " ".join(f'{k.replace("_", "-")}="{val}"' for k, val in attrs.items())
        return (f'<text {a} data-ref="{html.escape(ref)}"{y} data-sc="{sc}" data-fmt="{f}" data-v="{html.escape(repr(v))}">'
                f'<title>{html.escape(self.title(ref, year, sc))}</title>{fmt(v, f)}</text>')

    def src(self, *refs: str) -> str:
        """節末小字：Excel 位置（頁＋列 ID，具名範圍附註）。"""
        out = []
        for r in refs:
            if "!" in r:
                sheet, rid = r.split("!", 1)
                out.append(f"{sheet} {rid}")
                continue
            multi = ":" in self.loc.names.get(r, "")
            sheet, _, rid = self.loc.locate(r, YEARS[0] if multi else None)
            out.append(f"{sheet} {rid}（{r}）" if rid and rid != r else f"{sheet}（{r}）")
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
        lab = f"{MINUS if val < 0 else ''}{abs(val):,g}"
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
--inf:#2f855a;--rd:#9ae6b4;--con:#4a5568;--own:#a0aec0;--amz:#d69e2e;--oai:#8a94a0;--warn-bg:#fff4e5;--ok:#2f855a}
@media (prefers-color-scheme: dark){:root{--bg:#14181c;--fg:#e8ebee;--muted:#9aa4ae;--line:#323a42;--card:#1b2127;--accent:#5fc4cb;
--rev:#6aa7ea;--cost:#f08a4b;--gap:#ff7b72;--sub:#6aa7ea;--api:#3f7fbf;--ads:#9cc7ee;--net:#e8ebee;--tgt:#c792ea;
--inf:#56c288;--rd:#2e6b4a;--con:#a0aec0;--own:#5a6573;--amz:#e6b450;--oai:#6b7682;--warn-bg:#2b2416;--ok:#56c288}}
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
.lead{font-size:16px;font-weight:600;border-left:4px solid var(--amz);padding:6px 0 6px 12px;margin:4px 0 14px}
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
tr.sep td{border-top:1.5px solid var(--muted)}
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
.risk li{margin:8px 0}
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


def _xtick(o, x, y, H):
    o.append(f'<text class="tick" x="{x:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')


def chart_gw(V: Values) -> str:
    """每 VR 等值 GW 營收 vs 全成本（RMB 億／GW／年）。"""
    W, H, L, R, T, B = 640, 300, 52, 10, 26, 34
    rev = [V.get("COST_PropRev_VR", y) for y in YEARS]
    full = [V.get("COST_PropFull_VR", y) for y in YEARS]
    ymax = nice_max(max(full + rev) * 1.08)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="每 VR 等值 GW 營收與全成本 2025–2030">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "RMB 億/GW/年")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.3
    for i, y in enumerate(YEARS):
        x0 = L + slot * i + slot * 0.18
        for j, (vals, cls) in enumerate(((rev, "rev"), (full, "cost"))):
            yt = ymap(vals[i], H, T, B, 0, ymax)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
        yt = ymap(max(full[i], rev[i]), H, T, B, 0, ymax)
        o.append(V.t("COST_PropGap_VR", y, "f0", x=f"{x0 + bw:.1f}", y=f"{yt - 6:.1f}", text_anchor="middle", **{"class": "gaplbl"}))
        _xtick(o, L + slot * i + slot / 2, y, H)
    o.append("</svg>")
    return "\n".join(o)


def chart_rev(V: Values) -> str:
    """營收結構（截頂後）與分析師共識目標（RMB 億）。"""
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    tgt = [V.get("RVS_Target", y) for y in YEARS]
    tot = [V.get("Reverse!X16", y) for y in YEARS]
    ymax = nice_max(max(tgt + tot) * 1.08)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="營收結構與分析師共識">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "RMB 億")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.36
    for i, y in enumerate(YEARS):
        xc = L + slot * i + slot * 0.42
        base = 0.0
        for ref, cls in (("Reverse!X17", "api"), ("Reverse!X18", "sub"), ("Reverse!X19", "ads")):
            v = V.get(ref, y)
            if v <= 0:
                continue
            y1, y0 = ymap(base + v, H, T, B, 0, ymax), ymap(base, H, T, B, 0, ymax)
            o.append(f'<rect x="{xc - bw / 2:.1f}" y="{y1:.1f}" width="{bw:.1f}" height="{y0 - y1:.1f}" style="fill:var(--{cls})"/>')
            base += v
        o.append(V.t("Reverse!X16", y, "f1", x=f"{xc:.1f}", y=f"{ymap(base, H, T, B, 0, ymax) - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        if y >= 2026:
            yt = ymap(tgt[i], H, T, B, 0, ymax)
            xd = xc + bw / 2 + 8
            o.append(f'<path d="M{xd:.1f},{yt - 6:.1f} l6,6 l-6,6 l-6,-6 z" style="fill:var(--tgt)"/>')
            o.append(V.t("RVS_Target", y, "f0", x=f"{xd + 9:.1f}", y=f"{yt + 4:.1f}", **{"class": "lbl"}, style="fill:var(--tgt)"))
        _xtick(o, xc, y, H)
    o.append("</svg>")
    return "\n".join(o)


def chart_compute(V: Values) -> str:
    """實體 GW：推論（截頂後）＋研發＋閒置＝供給。"""
    W, H, L, R, T, B = 640, 300, 44, 10, 26, 34
    sup = [V.get("CMP_SupplyGW", y) for y in YEARS]
    ymax = nice_max(max(sup) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="算力供給 GW 的用途拆分">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "GW（實體，IT）")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.4
    for i, y in enumerate(YEARS):
        xc = L + slot * i + slot / 2
        base = 0.0
        for ref, cls in (("CMP_InfGW", "inf"), ("CMP_RDGW", "rd"), ("Compute!G106", "own")):
            try:
                v = V.get(ref, y)
            except ValueError:
                continue
            y1, y0 = ymap(base + v, H, T, B, 0, ymax), ymap(base, H, T, B, 0, ymax)
            if y0 - y1 > 0.05:
                o.append(f'<rect x="{xc - bw / 2:.1f}" y="{y1:.1f}" width="{bw:.1f}" height="{y0 - y1:.1f}" style="fill:var(--{cls})"/>')
            base += v
        o.append(V.t("CMP_SupplyGW", y, "f2", x=f"{xc:.1f}", y=f"{ymap(sup[i], H, T, B, 0, ymax) - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        _xtick(o, xc, y, H)
    o.append("</svg>")
    return "\n".join(o)


def chart_cash(V: Values) -> str:
    """年底現金：基準（可轉債現金償還）vs 轉股情境；橫線＝最低現金。"""
    W, H, L, R, T, B = 640, 280, 44, 10, 26, 34
    yrs = YEARS[1:]
    base = [V.get("FND_CashEnd", y) for y in yrs]
    conv = [V.get("FND_CashEndConv", y) for y in yrs]
    mins = [V.get("FND_MinCash", y) for y in yrs]
    ymax = nice_max(max(base + conv + mins) * 1.12)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="年底現金：基準與可轉債轉股情境">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "RMB 億")
    slot = (W - L - R) / len(yrs)
    bw = slot * 0.3
    for i, y in enumerate(yrs):
        x0 = L + slot * i + slot * 0.18
        for j, (ref, cls) in enumerate((("FND_CashEnd", "rev"), ("FND_CashEndConv", "amz"))):
            v = V.get(ref, y)
            yt = ymap(v, H, T, B, 0, ymax)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{yt:.1f}" width="{bw - 2:.1f}" height="{H - B - yt:.1f}" style="fill:var(--{cls})"/>')
            o.append(V.t(ref, y, "f0", x=f"{x0 + j * bw + bw / 2 - 1:.1f}", y=f"{yt - 5:.1f}", text_anchor="middle", **{"class": "lbl"}))
        ym = ymap(mins[i], H, T, B, 0, ymax)
        o.append(f'<line x1="{x0 - 4:.1f}" x2="{x0 + 2 * bw + 2:.1f}" y1="{ym:.1f}" y2="{ym:.1f}" style="stroke:var(--gap);stroke-width:2;stroke-dasharray:4 3"/>')
        _xtick(o, L + slot * i + slot / 2, y, H)
    o.append("</svg>")
    return "\n".join(o)


def chart_vs(V: Values) -> str:
    """每 VR 等值 GW 差額（$B）：智譜 vs OpenAI v0.6。"""
    W, H, L, R, T, B = 640, 280, 48, 10, 30, 34
    zp = [V.get("COST_PropGap_VR_USD", y) for y in YEARS]
    oa = [V.get("OAI_COST_PropGap_VR", y) for y in YEARS]
    ymin = -nice_max(-min(zp + oa) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="每 VR 等值 GW 差額：智譜與 OpenAI">']
    o += axis(W, H, L, R, T, B, ymin, 0, 4, "$B/GW/年")
    slot = (W - L - R) / len(YEARS)
    bw, gap = slot * 0.36, 8
    y0 = ymap(0, H, T, B, ymin, 0)
    for i, y in enumerate(YEARS):
        x0 = L + slot * i + (slot - 2 * bw - gap) / 2
        for j, (ref, cls, vals) in enumerate((("COST_PropGap_VR_USD", "gap", zp), ("OAI_COST_PropGap_VR", "oai", oa))):
            xb = x0 + j * (bw + gap)
            yb = ymap(vals[i], H, T, B, ymin, 0)
            o.append(f'<rect x="{xb:.1f}" y="{y0:.1f}" width="{bw:.1f}" height="{yb - y0:.1f}" style="fill:var(--{cls})"/>')
            o.append(V.t(ref, y, "f1", x=f"{xb + bw / 2:.1f}", y=f"{min(yb + 13, H - B - 2):.1f}", text_anchor="middle", style="font-size:10.5px"))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{T - 14}" text-anchor="middle">{y}</text>')
    o.append("</svg>")
    return "\n".join(o)


def year_table(V: Values, rows: list[tuple], years=YEARS, head="RMB 億") -> str:
    """rows：(標籤, ref, fmt, 樣式)。值為空的格顯示「—」。"""
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
    tmp = Path(tempfile.mkdtemp(prefix="z5_html_"))
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
    if V.get("CHK_Errors", None, "base") != 0:
        raise SystemExit("基準 CHK_Errors ≠ 0，不產生成品")

    stamp = model.name.split("_")[0]
    ver = re.search(r"_v([\d.]+)\.xlsx$", model.name).group(1)
    html_path = outdir / f"{stamp}_智譜收支模型_v{ver.replace('.', '_')}.html"
    xlsx_path = outdir / f"{stamp}_智譜收支模型_v{ver.replace('.', '_')}.xlsx"
    page = render(V, model, ver, stamp, cfg)
    outdir.mkdir(parents=True, exist_ok=True)
    html_path.write_text(page, encoding="utf-8")
    shutil.copyfile(model, xlsx_path)
    shutil.rmtree(tmp, ignore_errors=True)
    return html_path, xlsx_path, V


def latest_src_date(model: Path) -> str:
    ws = openpyxl.load_workbook(model, read_only=True)["SRC_ZP"]
    ds = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        d = row[10]                       # K＝文件日期
        if isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", d.strip()):
            ds.append(d.strip())
    return max(ds) if ds else "—"


def render(V: Values, model: Path, ver: str, stamp: str, cfg: dict) -> str:
    n = V.n
    Y = 2030
    H2 = "2H26"
    date_disp = f"{stamp[:4]}-{stamp[4:6]}-{stamp[6:]}"
    tk_file = V.get("TK_Link!B3")
    tk_ver = V.get("TK_Link!B4")
    tk_sha = V.get("TK_Link!B5")
    tk_date = V.get("TK_Link!B6")
    lab = {k: v["label"] for k, v in cfg["scenarios"].items()}

    S = []
    # ── 標頭 ──
    S.append(f"""<header><h1>智譜收支模型 v{ver}｜一頁摘要</h1>
<div class="meta">{date_disp}｜Excel <code>{html.escape(model.name)}</code>｜Tokenomics {html.escape(str(tk_ver))}（<code>{html.escape(str(tk_sha))[:7]}</code>）｜OpenAI v0.6 對照（OAI_Link）｜單位 RMB 億（人民幣億元）、與 OpenAI 並排時 $B，曆年制｜數字游標停留可見 Excel 位置；手機上圖表可左右滑動</div></header>""")

    # ── ① 命題與結論 ──
    S.append(f"""<section id="s1"><h2><span class="no">①</span>命題與一句話結論（兩種讀法）</h2>
<p class="note">命題：智譜（02513.HK）每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本？若不能，缺口要多少外部資金、由誰補？</p>
<p class="lead">不能，而且 2028 年後不再收斂。依招股章程更正算力費口徑（r2 P1）並設每 GW 算力租價下限（供應商不長期賠本出租）後，基準 2030 年每 VR 等值 GW 差額 {n("COST_PropGap_VR", Y, "f0")} RMB 億（{n("COST_PropGap_VR_USD", Y, "f1")} $B），覆蓋率 {n("COST_Coverage", Y, "pct0")}；2028–2030 差額 {n("COST_PropGap_VR", 2028, "f0")}／{n("COST_PropGap_VR", 2029, "f0")}／{n("COST_PropGap_VR", 2030, "f0")}。資金面，模型基準 2030 前累計外部資金需求只有 {n("FND_ExtNeedCum", Y, "f1")} 億（{n("FND_FirstGapYear", None, "yrt")} 年補足最低現金），但這個結論可信度低（見下）。</p>
<div class="kpis">
<div class="kpi"><div class="k">讀法一：2030 每 VR 等值 GW 差額（命題定義）</div><div class="v neg">{n("COST_PropGap_VR", Y, "f0")}</div><div class="s">RMB 億/GW/年；＝{n("COST_PropGap_VR_USD", Y, "f1")} $B</div></div>
<div class="kpi"><div class="k">讀法二：2030 每實體 GW 差額</div><div class="v neg">{n("COST_PropGap_Phys_USD", Y, "f2")}</div><div class="s">$B/GW/年；＝{n("COST_PropGap_Phys", Y, "f1")} RMB 億；2025 為 {n("COST_PropGap_Phys_USD", 2025, "f1")}</div></div>
<div class="kpi"><div class="k">2030 覆蓋率（營收 ÷ 全成本）</div><div class="v">{n("COST_Coverage", Y, "pct0")}</div><div class="s">2025 {n("COST_Coverage", 2025, "pct0")}、2028 {n("COST_Coverage", 2028, "pct0")}；全成本含股權報酬</div></div>
<div class="kpi"><div class="k">模型基準：累計外部資金需求 2026–2030</div><div class="v">{n("FND_ExtNeedCum", Y, "f1")}</div><div class="s">RMB 億；首次缺口年 {n("FND_FirstGapYear", None, "yrt")}；<b>可信度低</b>：實際燒錢速度是模型的 {n("FND_UseVsModel", H2, "f0")} 倍</div></div>
</div>
<h3>兩種讀法怎麼看</h3>
<p><b>讀法一（每 VR 等值 GW）</b>是命題定義，與 OpenAI 同口徑：把算力換成 NVIDIA VR200 的產能再除。<b>讀法二（每實體 GW）</b>直接除以智譜實際租用的 GW。白話說：<b>智譜的國產／H20／Hopper 機隊換成 VR200 等值，只值約 {n("Cost!K67", Y, "f3")} 倍</b>（2030；2026 為 {n("Cost!K67", 2026, "f3")}），所以讀法一的分母很小、每 GW 數字被放大。2030 每實體 GW 營收 {n("COST_PropRev_Phys_USD", Y, "f2")} $B、全成本 {n("COST_PropFull_Phys_USD", Y, "f2")} $B——不要把讀法一的「每 GW 數千億人民幣」理解成智譜每 GW 真的燒那麼多錢；它燒得少，是因為它的 GW 很便宜也很弱。</p>
<div class="must"><b>命題 2「2030 前大致不需外部資金」可信度低。</b>這是模型基準；但唯一即時觀察——2026-07 配售款 7–8 月已動用 {n("Funding!F52", H2, "f1")} 億，模型同期只流出 {n("Funding!F53", H2, "f1")} 億（{n("FND_UseVsModel", H2, "f1")} 倍）——與 1H26 現金對帳差 {n("Funding!F26", "1H26", "f1")} 億，都指向模型<b>低估燒錢</b>。反方向的落差也要對稱看：模型 2H26 年化雲端營收對公司 MaaS ARR 差距 {n("REV_ARRGapMaaS", None, "pct0")}，指向<b>營收低估</b>。兩者都未證實（第 ⑦ 節）。</div>
<div class="must"><b>翻轉點（r2：招股章程 P1 後重算；每一點都在該值以 LibreOffice 重算 Excel 確認）</b>
<ul>
<li><b>命題 1（2030 每 VR 等值 GW 轉正）只剩一個翻轉點</b>：研發項下算力費占研發開支降到 ≤ {n("INP_137", None, "pct1", "flip_rdfee")}（基準 {n("INP_137", None, "pct1")}＝招股章程 1H25 實際），也就是 1H26 總算力費只有 {n("Compute!G75", "1H26", "f1", "flip_rdfee")} 億、不是 {n("Compute!G75", "1H26", "f1")} 億——此時 2030 差額 {n("COST_PropGap_VR", Y, "f0", "flip_rdfee")}。白話：未來算力成本是按 1H26「每 token 花多少算力費」外推的，只有這個比例被高估一半以上，2030 才會打平。</li>
<li><b>不是翻轉點</b>：每 GW 價格年降幅（租價有下限；降 25% 時 2030 差額仍為 {n("COST_PropGap_VR", Y, "f0", "px_m25")}）、供應商最低毛利（搜尋到 −90% 仍不轉正）、研發占比路徑、2H26 任務數成長（取高 +{n("INP_038", None, "pct0", "task_hi")} 時 2030 差額 {n("COST_PropGap_VR", Y, "f0", "task_hi")}，搜尋到 +2,000% 仍不轉正）。</li>
<li><b>命題 2（2030 前不需外部資金）</b>：研發項下算力費占比 ≤ {n("INP_137", None, "pct1", "flip_rdfee_ext")}、或營業成本計算服務費占 API 銷售成本 ≤ {n("INP_102", None, "pct0", "flip_inffee_ext")}（基準 {n("INP_102", None, "pct0")}）、或供應商最低毛利 ≤ {n("INP_136", None, "pct1", "flip_mgm_ext")}（供應商願意小幅賠本出租）；在這些值下累計外部資金需求分別為 {n("FND_ExtNeedCum", Y, "f1", "flip_rdfee_ext")}／{n("FND_ExtNeedCum", Y, "f1", "flip_inffee_ext")}／{n("FND_ExtNeedCum", Y, "f1", "flip_mgm_ext")}。</li>
</ul>
Coding Plan 額度使用率與 Coding Plan 輸出占比不列為翻轉點：它們同時改變 1H26 的 η 校準，結果是假象而非經濟機制（第 ⑥ 節註明）。招股章程更正後，Z5b 的「研發占比 2030 ≤0.572」不再成立：研發占比起點已是訓練最低占比 0.30，無法再降。翻轉值由 <code>tools/sensitivity_z4.py</code>（設定 <code>tools/sensitivity_r2.yaml</code>）以 engine 執行 Excel 二分搜尋得到。</div>
{V.src("COST_PropGap_VR", "COST_PropGap_VR_USD", "COST_PropGap_Phys", "COST_PropGap_Phys_USD", "Cost!K67", "COST_Coverage", "FND_ExtNeedCum", "FND_FirstGapYear", "FND_UseVsModel", "Funding!F52", "Funding!F53", "Funding!F26", "REV_ARRGapMaaS", "INP_137", "INP_102", "INP_038", "INP_136", "Compute!G75")}
</section>""")

    # ── ② 每 VR 等值 GW ──
    S.append(f"""<section id="s2"><h2><span class="no">②</span>每 GW：營收 vs 全成本（2025–2030）</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>營收淨額（截頂後）</span><span><i style="background:var(--cost)"></i>全成本（算力＋本地化交付＋非算力＋股權報酬，現金口徑）</span><span><b style="color:var(--gap)">紅字</b>＝每 VR 等值 GW 差額</span></div>
<div class="cw">{chart_gw(V)}</div>
{year_table(V, [("營收淨額", "COST_PropRev_VR", "f0", ""), ("　算力成本（算力服務費實付／租用）", "COST_PropCompute_VR", "f0", ""),
                ("　本地化部署交付成本", "COST_PropOnPrem_VR", "f0", ""), ("　非算力成本（不含股權報酬）", "COST_PropNonComp_VR", "f0", ""),
                ("　股權報酬", "COST_PropSBC_VR", "f0", ""), ("全成本", "COST_PropFull_VR", "f0", ""),
                ("差額", "COST_PropGap_VR", "f0", "em neg"), ("覆蓋率", "COST_Coverage", "pct0", ""),
                ("分母：供給 VR 等值 GW", "CMP_Supply_VReq", "f3", "")], head="每 VR 等值 GW（RMB 億/GW/年）")}
{year_table(V, [("分母：供給 GW（實體）", "Cost!K66", "f3", ""), ("機隊 VR 等值係數", "Cost!K67", "f3", ""),
                ("每實體 GW 營收淨額", "COST_PropRev_Phys", "f0", ""), ("每實體 GW 全成本", "COST_PropFull_Phys", "f0", ""),
                ("每實體 GW 差額", "COST_PropGap_Phys", "f1", "em neg"), ("每實體 GW 差額（$B）", "COST_PropGap_Phys_USD", "f2", "")], head="每實體 GW（RMB 億/GW/年）")}
{year_table(V, [("營收淨額", "Cost!K35", "f1", ""), ("全成本（含股權報酬）", "COST_FullCash", "f1", ""), ("差額", "COST_GapCash", "f1", "em neg")], head="絕對金額（RMB 億）")}
<p class="note">2027 年前差額快速收斂（營收成長快於成本）；2028 年起每 GW 租價觸及下限＝供應商持有成本（組合後 {n("CMP_PricePerGW", 2028, "f0")}→{n("CMP_PricePerGW", Y, "f0")}，未設下限時會降到 {n("Compute!G176", Y, "f0")}），每 GW 算力成本不再下降，差額不再收斂（2028–2030 見上表）。2025 分母只有 {n("CMP_Supply_VReq", 2025, "f3")} VR 等值 GW，每 GW 數字特別大。絕對金額：2030 年差額 {n("COST_GapCash", Y, "f1")} 億。</p>
{V.src("COST_PropRev_VR", "COST_PropCompute_VR", "COST_PropOnPrem_VR", "COST_PropNonComp_VR", "COST_PropSBC_VR", "COST_PropFull_VR", "COST_PropGap_VR", "COST_Coverage", "CMP_Supply_VReq", "Cost!K66", "Cost!K67", "COST_PropRev_Phys", "COST_PropFull_Phys", "COST_PropGap_Phys", "COST_PropGap_Phys_USD", "Cost!K35", "COST_FullCash", "COST_GapCash")}
</section>""")

    # ── ③ 營收結構 ──
    S.append(f"""<section id="s3"><h2><span class="no">③</span>營收結構與分析師共識</h2>
<div class="legend"><span><i style="background:var(--api)"></i>API 按量計費</span><span><i style="background:var(--sub)"></i>Coding Plan 訂閱</span><span><i style="background:var(--ads)"></i>本地化部署＋其他</span><span><i style="background:var(--tgt);transform:rotate(45deg)"></i>分析師共識平均（只對照）</span></div>
<div class="cw">{chart_rev(V)}</div>
{year_table(V, [("API 按量計費（截頂後）", "Reverse!X17", "f1", ""), ("Coding Plan 訂閱（截頂後）", "Reverse!X18", "f1", ""), ("本地化部署＋其他（不截頂）", "Reverse!X19", "f1", ""),
                ("營收淨額（正向）", "Reverse!X16", "f1", "em"), ("共識平均（3 筆；2029–30 持平）", "RVS_Target", "f1", ""),
                ("共識最低", "RVS_TargetLo", "f1", ""), ("共識最高", "RVS_TargetHi", "f1", ""), ("共識 ÷ 正向", "RVS_Ratio", "f2", "")])}
<p class="note">管理層沒有量化的營收或獲利目標（Reverse X01：不存在），只有 2026 年末 ARR 指引 {n("Reverse!X02", H2, "f1")} 億（模型 2H26 年化 {n("Reverse!X03", H2, "f1")}，差距 {n("Reverse!X04", H2, "pct0")}；只對照、不反推）。共識 2028 是正向的 {n("RVS_Ratio", 2028, "f2")} 倍；只靠 API 補足需 {n("RVS_MultAPI", 2028, "f2")} 倍、只靠 Coding Plan 需 {n("RVS_MultSub", 2028, "f2")} 倍。即使共識全數實現、支出不變，反向累計外部資金需求 {n("RVS_ExtNeedCum", Y, "f0")}（反向模式，不回饋基準）。</p>
{V.src("Reverse!X16", "Reverse!X17", "Reverse!X18", "Reverse!X19", "RVS_Target", "RVS_TargetLo", "RVS_TargetHi", "RVS_Ratio", "Reverse!X02", "Reverse!X03", "Reverse!X04", "RVS_MultAPI", "RVS_MultSub", "RVS_ExtNeedCum")}
</section>""")

    # ── ④ 算力 ──
    S.append(f"""<section id="s4"><h2><span class="no">④</span>算力：供給 GW 的用途（實體 GW，IT 電力）</h2>
<div class="legend"><span><i style="background:var(--inf)"></i>推論 GW（截頂後）</span><span><i style="background:var(--rd)"></i>研發 GW</span><span><i style="background:var(--own)"></i>閒置保留（2H26 起）</span><span>柱頂數字＝供給 GW</span></div>
<div class="cw">{chart_compute(V)}</div>
{year_table(V, [("推論 GW（截頂後）", "CMP_InfGW", "f3", ""), ("研發 GW", "CMP_RDGW", "f3", ""), ("閒置 GW", "Compute!G106", "f3", ""),
                ("供給 GW（租用＋自有）", "CMP_SupplyGW", "f3", "em"), ("其中自有", "CMP_Supply_Owned", "f3", ""),
                ("機隊 VR 等值係數", "CMP_VReqFactor", "f3", ""), ("供給 VR 等值 GW（命題分母）", "CMP_Supply_VReq", "f3", ""),
                ("η（P1：基準 1；1H26 殘差不足而上調，之後沿用）", "CMP_Eta", "f2", ""), ("研發占非閒置供給（殘差）", "CMP_RDShare", "pct0", ""),
                ("每 GW 年租價（組合後；2H26 起不低於持有成本）", "CMP_PricePerGW", "f0", "em"), ("　對照：未設下限的觀察租價", "Compute!G176", "f0", ""),
                ("　供應商持有成本（組合後＝下限）", "COST_HoldW", "f0", ""), ("　供應商推算毛利率", "COST_CloudGMPct", "pct0", "")], head="GW／比例／RMB 億")}
<p><b>招股章程更正（r2 P1）</b>：研發項下的算力費也涵蓋推理（第 255 頁），營業成本中的計算服務費只占總算力費 2–3%（1H25 {n("Compute!G192", 2025, "pct1")}，Compute G192）。所以推論與研發不再以會計科目拆分：供給 GW＝總算力費（1H26 {n("Compute!G75", "1H26", "f1")} 億＝營業成本計算服務費 {n("Compute!G68", "1H26", "f2")}＋研發開支 × {n("INP_137", None, "pct1")}）÷ 每 GW 價格；推論 GW＝token 換算 GW ÷ η，η 基準＝1（信任 Tokenomics 物理）；研發 GW＝殘差（同 OpenAI）。1H26 token 換算 GW {n("Compute!G55", "1H26", "f3")} 已是供給 {n("Compute!G104", "1H26", "f3")} 的兩倍多，所以 η 只能上調到 <b>{n("CMP_Eta1H26", None, "f2")}</b>（使研發殘差＝訓練最低占比 {n("INP_109", None, "pct0")}；Checks WARN），2025 為 {n("CMP_Eta2025", None, "f2")}；舊法（API 銷售成本＝推論支出）的 η 是 {n("Compute!G198", "1H26", "f2")}。2H26 起 η 與研發占比沿用 1H26，供給 GW 由需求配置：有效推論 ÷［(1−閒置)(1−研發占比)］。因為未來供給只依 1H26「供給 ÷ token」比例外推，<b>新舊兩法的命題結果相同</b>（舊法情境 2030 差額 {n("COST_PropGap_VR", Y, "f0", "eta_old")}）；P1 改的是推論與研發的拆分、η 的大小，以及總算力費的估計（研發開支 × {n("INP_137", None, "pct1")}；Z3 為扣股權報酬後 × {n("INP_103", None, "pct1")}）。</p>
<p class="note">租價下限（Z5b V1）：觀察租價每年降 10% 會在 2027–2028 年低於供應商持有成本（原模型 2028 起供應商毛利率 −9%～−37%，即長期賠本出租，不可持續）；現在 2H26 起每 GW 租價＝MAX（觀察租價, 持有成本 ×（1＋最低毛利 {n("INP_136", None, "pct0")}））。</p>
{V.src("CMP_InfGW", "CMP_RDGW", "Compute!G106", "CMP_SupplyGW", "CMP_Supply_Owned", "CMP_VReqFactor", "CMP_Supply_VReq", "CMP_Eta", "CMP_RDShare", "CMP_PricePerGW", "Compute!G176", "COST_HoldW", "COST_CloudGMPct", "INP_136", "Compute!G192", "Compute!G75", "Compute!G68", "INP_137", "Compute!G55", "Compute!G104", "CMP_Eta1H26", "INP_109", "CMP_Eta2025", "Compute!G198")}
</section>""")

    # ── ⑤ 現金與外部資金 ──
    S.append(f"""<section id="s5"><h2><span class="no">⑤</span>現金與外部資金需求</h2>
<div class="legend"><span><i style="background:var(--rev)"></i>年底現金（基準：可換股債券 2027-09 現金償還）</span><span><i style="background:var(--amz)"></i>情境：可換股債券轉股</span><span><i style="background:var(--gap);height:3px"></i>最低現金（次年成本 6 個月）</span></div>
<div class="cw">{chart_cash(V)}</div>
{year_table(V, [("營收淨額", "Funding!F01", "f1", ""), ("算力成本（現金）", "Funding!F02", "f1", ""), ("本地化部署交付成本", "Funding!F03", "f1", ""),
                ("非算力成本（不含股權報酬）", "Funding!F04", "f1", ""), ("自由現金流（股權報酬加回、＋其他收益、−利息）", "FND_FCF", "f1", "em"),
                ("融資前淨現金流（−資本支出、−併購）", "FND_NetOp", "f1", ""), ("已到位融資", "FND_Committed", "f1", ""),
                ("可換股債券償還", "FND_DebtRepay", "f1", ""), ("最低現金", "FND_MinCash", "f1", ""),
                ("當年外部資金需求", "FND_ExtNeed", "f1", "em"), ("累計外部資金需求", "FND_ExtNeedCum", "f1", "em"),
                ("年底現金", "FND_CashEnd", "f1", ""), ("轉股情境：累計外部資金需求", "FND_ExtNeedCumConv", "f1", ""),
                ("轉股情境：年底現金", "FND_CashEndConv", "f1", "")])}
<p class="note">來源順序：①期初現金（2025 年底 {n("FND_CashEnd", 2025, "f1")}）→ ②已到位股權：2026-01 IPO {n("Funding!F14", "1H26", "f1")}（1H26 已全數動用）、2026-07 配售 {n("Funding!F15", H2, "f1")}、2026-09 配售 {n("Funding!F16", H2, "f1")} → ③債務：可換股債券 {n("Funding!F18", H2, "f1")}（2027-09 現金償還 {n("FND_DebtRepay", 2027, "f1")}）、銀行借款 {n("FND_BankLoans", Y, "f2")} 續借 → ④外部資金。補足者是港股公開市場投資人（兩次折價配售）與可換股債券投資人。<b>只靠 IPO＋2026-07 配售</b>：2030 年底現金 {n("FND_CashEnd", Y, "f1", "only_pl1")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f1", "only_pl1")}；<b>只靠 IPO</b>：首次缺口年 {n("FND_FirstGapYear", None, "yrt", "ipo_only")}、2030 累計 {n("FND_ExtNeedCum", Y, "f1", "ipo_only")}。轉股稀釋 {n("Funding!F45", None, "pct1")}。或有：可換股債券本金 {n("FND_Contingent", None, "f1")}；租賃負債 {n("FND_LeaseLiab", None, "f2")} 億（2025 年底；租賃算力硬件與辦公室，含 2024 年售後租回 {n("Funding!F89", None, "f2")} 億；利率 {n("Funding!F88", None, "pct2")}）列債務對照——租賃付款不在算力服務費內，但其折舊已在校準期非算力成本內，不另加（r2 P3）；政府補助沖減開支、已含在費用淨額（r2 P4）；未提用授信、算力採購承諾找不到（Funding F49）。</p>
{V.src("Funding!F01", "Funding!F02", "Funding!F03", "Funding!F04", "FND_FCF", "FND_NetOp", "FND_Committed", "FND_DebtRepay", "FND_MinCash", "FND_ExtNeed", "FND_ExtNeedCum", "FND_CashEnd", "FND_ExtNeedCumConv", "FND_CashEndConv", "Funding!F14", "Funding!F15", "Funding!F16", "Funding!F18", "FND_BankLoans", "FND_FirstGapYear", "Funding!F45", "FND_Contingent", "FND_LeaseLiab", "Funding!F88", "Funding!F89")}
</section>""")

    # ── ⑥ 關鍵驅動與敏感度 ──
    rows = ["base", "rdfee_lo", "rdfee_hi", "inffee_lo", "px_m25", "px_0", "mgm_hi", "eta_conv", "eta_old", "rd_05", "util_hi", "util_lo", "task_lo", "task_hi", "dc1gw", "stress"]
    tr = []
    for sc in rows:
        tr.append(f'<tr class="{"em" if sc == "base" else ""}"><td>{html.escape(lab[sc])}</td><td>{n("REV_NetCapped", Y, "f1", sc)}</td>'
                  f'<td>{n("COST_PropGap_VR", Y, "f0", sc)}</td><td>{n("COST_Coverage", Y, "f2", sc)}</td><td>{n("CMP_SupplyGW", Y, "f2", sc)}</td>'
                  f'<td>{n("FND_ExtNeedCum", Y, "f0", sc)}</td><td>{n("FND_FirstGapYear", None, "yrt", sc)}</td><td>{n("FND_CashEnd", Y, "f0", sc)}</td></tr>')
    S.append(f"""<section id="s6"><h2><span class="no">⑥</span>三個關鍵驅動與敏感度</h2>
<p class="note">r2（招股章程 P1 後）的三個關鍵驅動：①總算力費的估計——研發項下算力費占研發開支（INP_137）與營業成本計算服務費占 API 銷售成本（INP_102），決定 1H26「每 token 花多少算力費」，未來算力成本依此外推 ②每 GW 算力價格：年變動（INP_100）與租價下限的最低毛利（INP_136）③η 路徑（INP_105：沿用 1H26 上調值或收斂至 1）。另列 η 舊法、研發占比路徑、Coding Plan 使用率、2H26 任務成長、1 GW 資料中心與壓力組合。設定取自 Inputs 低／高欄；全部由 Excel 重算。2030 差額單位 RMB 億／VR 等值 GW。</p>
<div class="tw"><table><thead><tr><th>情境</th><th>2030 營收淨額</th><th>2030 每 VR GW 差額</th><th>2030 覆蓋率</th><th>2030 供給 GW</th><th>累計外部資金需求 2030</th><th>首次缺口年</th><th>2030 年底現金</th></tr></thead><tbody>
{"".join(tr)}
</tbody></table></div>
<p class="note"><b>Coding Plan 額度使用率（與輸出占比）的擺幅是假象</b>：它們同時改變 1H26 的 token 換算 GW 與 η 校準（η 吸收），2H26 起才因 token 結構不同而分歧；使用率 {n("INP_060", None, "f3", "flip_util")} 時 2030 差額 {n("COST_PropGap_VR", Y, "f0", "flip_util")}，但這不是經濟機制，因此不列為翻轉點。η 舊法與基準結果相同（第 ④ 節）；研發占比升至 0.5 使供給與成本上升。每 GW 價格降 25% 時 2030 差額與基準相同（租價下限），只改 2026–2027 的現金。</p>
<div class="must"><b>命題 2</b>：模型基準只在 2030 年為補足最低現金需要 {n("FND_ExtNeedCum", Y, "f1")} 億（年底現金 {n("FND_CashEnd", Y, "f0")}，不是現金歸零）；最低毛利取高 → {n("FND_ExtNeedCum", Y, "f0", "mgm_hi")}（{n("FND_FirstGapYear", None, "yrt", "mgm_hi")}）；每 GW 價格不變 → {n("FND_ExtNeedCum", Y, "f0", "px_0")}（{n("FND_FirstGapYear", None, "yrt", "px_0")}）；η 收斂至 1 → {n("FND_ExtNeedCum", Y, "f0", "eta_conv")}（{n("FND_FirstGapYear", None, "yrt", "eta_conv")}；招股章程更正後 η 起點降低，此情境比 Z5b 小得多）；研發項下算力費占比取高 → {n("FND_ExtNeedCum", Y, "f0", "rdfee_hi")}、取低 → {n("FND_ExtNeedCum", Y, "f0", "rdfee_lo")}；1 GW 資料中心 → {n("FND_ExtNeedCum", Y, "f0", "dc1gw")}（{n("FND_FirstGapYear", None, "yrt", "dc1gw")}）。</div>
{V.src("REV_NetCapped", "COST_PropGap_VR", "COST_Coverage", "CMP_SupplyGW", "FND_ExtNeedCum", "FND_FirstGapYear", "FND_CashEnd", "INP_060", "INP_136", "INP_137", "INP_102")}
</section>""")

    # ── ⑦ 主要風險／與實際觀察的落差 ──
    S.append(f"""<section id="s7"><h2><span class="no">⑦</span>主要風險／與實際觀察的落差</h2>
<p class="note">以下是模型與公司實際揭露對不上的地方；任何一項若證實，結論（特別是命題 2）可能改變。支出面的落差（第 1、4 項）指向燒錢被低估，營收面的落差（第 5 項）指向營收被低估，方向相反、都未證實。</p>
<ol class="risk">
<li><b>配售款動用速度遠高於模型</b>：2026-07 配售款截至 2026-08-31 已動用 HK$ {n("SRC_ZP_540", None, "hm2y")} 億（SRC_ZP_540；約 RMB {n("Funding!F52", H2, "f1")} 億），是模型同期（7–8 月）融資前淨現金流出 {n("Funding!F53", H2, "f1")} 億的 <b>{n("FND_UseVsModel", H2, "f1")} 倍</b>（Funding F52–F54）。這是唯一的即時觀察，與 1H26 現金對帳差（第 4 項）同樣指向<b>模型低估燒錢</b>；若這筆是算力預付或 1 GW 資料中心支出，命題 2「2030 前大致不需外部資金」不成立：1 GW 資料中心情境下 2030 累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "dc1gw")} 億、首次缺口年 {n("FND_FirstGapYear", None, "yrt", "dc1gw")}。市值隱含營收的淨現金已扣除這筆已動用款（第 ⑧ 節，V3）。</li>
<li><b>公司說法前後不一：資料中心</b>（r2 P5）。招股章程（2025-12-30，第 223 頁）稱「並無計劃開發我們自身的 AI 數據中心」，列入實體清單後未採購任何 AI 晶片、全部向雲服務商採購算力（SRC_ZP_810、SRC_ZP_1051）；七個月後（2026-07）公司卻稱「已落地 1GW 級國產 AI 算力數據中心」（SRC_ZP_458：{n("SRC_ZP_458", None, "f0")} GW，口徑未明；若為設施口徑，÷ PUE ＝ IT {n("Compute!G130", H2, "f2")} GW，Compute G130）；模型基準 2H26 供給只有 {n("Compute!G104", H2, "f2")} GW（實體，Compute G104），1 GW 只列情境（INP_112）：2027 計入資本支出、2028 投產（2030 供給 {n("CMP_SupplyGW", Y, "f2", "dc1gw")} GW，其中自有 {n("CMP_Supply_Owned", Y, "f2", "dc1gw")} GW），投產後計入營運費用（TK IF_OpexGW，2030 {n("CMP_OwnedOpex", Y, "f1", "dc1gw")} 億）。<b>現金口徑</b>：2027 每 VR 等值 GW 差額 {n("COST_PropGap_VR", 2027, "f0", "dc1gw")}（一次性資本支出）、2030 {n("COST_PropGap_VR", Y, "f0", "dc1gw")}；<b>攤提口徑</b>（資本支出依 TK 折舊年限 {n("Compute!G180", Y, "f0", "dc1gw")} 年攤提，Cost 第十節）：2030 差額 {n("COST_PropGap_VR_Amort", Y, "f0", "dc1gw")}、覆蓋率 {n("COST_Coverage_Amort", Y, "pct0", "dc1gw")}——投產後並未轉正；累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "dc1gw")}、2030 年底現金 {n("FND_CashEnd", Y, "f0", "dc1gw")}。</li>
<li><b>國產晶片同口徑對照</b>：公司稱推論用國產晶片 10 萬張級（SRC_ZP_452），換算 IT {n("Compute!G124", H2, "f2")} GW；同口徑的模型 2H26 <b>國產推論 GW</b> 為 {n("Compute!G127", H2, "f3")} GW（差距 {n("Compute!G129", H2, "pct0")}，Compute G127、G129；招股章程 P1 後推論占供給七成，差距縮小）。供給口徑另列：模型 2H26 國產供給 {n("Compute!G126", H2, "f3")} GW（差距 {n("Compute!G128", H2, "pct0")}）。</li>
<li><b>預付算力服務費暴增</b>：2025 年底 {n("SRC_ZP_220", None, "f2")} 億 → 2026-06-30 {n("SRC_ZP_219", None, "f2")} 億（SRC_ZP_220／219，含其他）。模型算力成本＝算力服務費實付（費用口徑），預付不在內；1H26 現金對帳差 {n("Funding!F26", "1H26", "f2")} 億（Funding F26）可能部分來自此。</li>
<li><b>營收低於公司 ARR</b>：模型 2H26 年化雲端營收 {n("Revenue!R70", H2, "f1")} 億，公司 MaaS ARR（2026-08 月度年化）{n("Revenue!R71", H2, "f1")} 億，差距 {n("REV_ARRGapMaaS", None, "pct0")}（Revenue R70–R72；Checks WARN；不反推）。若 ARR 為真，營收被低估、命題 1 偏保守（與第 1、4 項方向相反）。</li>
<li><b>η 仍約 {n("CMP_Eta1H26", None, "f1")}</b>（Compute G94；2025 {n("CMP_Eta2025", None, "f2")}）：招股章程更正（P1）後已不是會計科目拆分造成，而是 1H26 的 token 量以 Tokenomics 物理換算，已超過總算力費買得到的 GW——國產晶片實際效率、token 量估計（API 由營收倒推、Coding Plan 用量）或租價三者至少一個有偏差，未證實。若 η 收斂至 1（INP_105），2030 供給 {n("CMP_SupplyGW", Y, "f2", "eta_conv")} GW、2030 差額 {n("COST_PropGap_VR", Y, "f0", "eta_conv")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "eta_conv")}——是命題 1、2 最大的單一不確定。</li>
</ol>
<p class="note">已在 Excel 的兩個情境：<b>1 GW 資料中心</b>（見上，INP_112）；<b>可換股債券轉股</b>（Funding 第五節，不受開關影響）：2030 年底現金 {n("FND_CashEndConv", Y, "f1")}（基準 {n("FND_CashEnd", Y, "f1")}）、累計外部資金需求 {n("FND_ExtNeedCumConv", Y, "f0")}、稀釋 {n("Funding!F45", None, "pct1")}；現價 ÷ 換股價 {n("Funding!F47", None, "f2")}。壓力組合（η 收斂＋1 GW＋價格不變＋員工年增取高）累計 {n("FND_ExtNeedCum", Y, "f0", "stress")}。</p>
{V.src("SRC_ZP_540", "Funding!F52", "Funding!F53", "FND_UseVsModel", "SRC_ZP_458", "Compute!G130", "Compute!G104", "Compute!G124", "Compute!G126", "Compute!G127", "Compute!G128", "Compute!G129", "CMP_OwnedOpex", "Compute!G180", "COST_PropGap_VR_Amort", "COST_Coverage_Amort", "SRC_ZP_452", "SRC_ZP_810", "SRC_ZP_1051", "INP_112", "SRC_ZP_220", "SRC_ZP_219", "Funding!F26", "Revenue!R70", "Revenue!R71", "REV_ARRGapMaaS", "CMP_Eta1H26", "CMP_Eta2025", "INP_105", "FND_CashEndConv", "FND_ExtNeedCumConv", "Funding!F45", "Funding!F47")}
</section>""")

    # ── ⑧ 市值對照 ──
    fw = "".join(f"<td>{n('RVS_FwdOverImplied', y, 'f2')}</td>" for y in YEARS)
    rv = "".join(f"<td>{n('Reverse!X16', y, 'f1')}</td>" for y in YEARS)
    hd = "".join(f"<th>{y}</th>" for y in YEARS)
    S.append(f"""<section id="s8"><h2><span class="no">⑧</span>市值對照：市值隱含營收 vs 正向營收</h2>
<div class="kpis">
<div class="kpi"><div class="k">市值（2026-10-07 收盤）</div><div class="v">{n("RVS_MktCap", None, "f0")}</div><div class="s">RMB 億；HK$ {n("Reverse!X38", None, "f0")}/股 × {n("Reverse!X39", None, "f0")} 股</div></div>
<div class="kpi"><div class="k">企業價值 EV＝市值 − 淨現金</div><div class="v">{n("RVS_EV", None, "f0")}</div><div class="s">RMB 億；淨現金 {n("RVS_NetCash", None, "f1")}（含 7、9 月新資金，扣 7–8 月已動用 {n("Funding!F52", H2, "f1")}）</div></div>
<div class="kpi"><div class="k">市值隱含營收＝EV ÷ 同業倍數</div><div class="v">{n("RVS_MktImpliedRev", None, "f1")}</div><div class="s">RMB 億；區間 {n("RVS_MktImpliedRevLo", None, "f1")}–{n("RVS_MktImpliedRevHi", None, "f1")}（＝{n("Reverse!X56", None, "f2")} $B）</div></div>
<div class="kpi"><div class="k">智譜自身 EV ÷ 1H26 年化營收</div><div class="v">{n("Reverse!X57", None, "f0")} 倍</div><div class="s">÷ 模型 2H26 年化：{n("Reverse!X58", None, "f1")} 倍；同業 MiniMax {n("RVS_PeerMult", None, "f1")} 倍</div></div>
</div>
<div class="tw"><table><thead><tr><th>RMB 億／倍</th>{hd}</tr></thead><tbody>
<tr><td>正向營收淨額</td>{rv}</tr>
<tr class="em"><td>正向營收 ÷ 市值隱含營收（正向營收倍數）</td>{fw}</tr>
</tbody></table></div>
<p class="note">同業倍數＝MiniMax EV ÷ 年化營收 {n("RVS_PeerMult", None, "f1")} 倍（區間 {n("Reverse!X51", None, "f1")}–{n("Reverse!X52", None, "f1")}，Inputs 係數）。以此倍數計，智譜市值「要求」約 {n("RVS_MktImpliedRev", None, "f0")} 億年營收；正向 {n("RVS_MktImpliedYear", None, "yrt")} 年就達到、2030 是它的 {n("RVS_FwdOverImplied", Y, "f1")} 倍；共識 2028 是它的 {n("Reverse!X62", 2028, "f1")} 倍。也就是說，<b>市值的前提不是比正向更高的營收，而是與同業同樣高的倍數</b>；用 1H26 實際營收算，智譜自己的 EV／營收是 {n("Reverse!X57", None, "f0")} 倍。股數含非上市股份（只計 H 股則市值約減半）。淨現金已扣除 2026-07 配售款截至 8 月底已動用的部分（Z5b V3；未扣 9 月以後的營運消耗）。</p>
{V.src("RVS_MktCap", "Reverse!X38", "Reverse!X39", "RVS_NetCash", "RVS_EV", "RVS_PeerMult", "Reverse!X51", "Reverse!X52", "RVS_MktImpliedRev", "RVS_MktImpliedRevLo", "RVS_MktImpliedRevHi", "Reverse!X56", "Reverse!X57", "Reverse!X58", "RVS_FwdOverImplied", "RVS_MktImpliedYear", "Reverse!X62")}
</section>""")

    # ── ⑨ 與 OpenAI v0.6 並排 ──
    pairs = [("覆蓋率（主讀）", "COST_Coverage", "OAI_COST_Coverage", "f2", "em"),
             ("自由現金流（$B）", "FND_FCF_USD", "OAI_FND_FCF", "f2", ""),
             ("已到位融資（$B）", "FND_Committed_USD", "OAI_FND_Committed", "f2", ""),
             ("累計外部資金需求（$B）", "FND_ExtNeedCum_USD", "OAI_FND_ExtNeedCum", "f1", "em"),
             ("年底現金（$B）", "FND_CashEnd_USD", "OAI_FND_CashEnd", "f2", ""),
             ("供給 VR 等值 GW", "CMP_Supply_VReq", "OAI_CMP_Supply_VReq", "f3", ""),
             ("每 VR 等值 GW 營收（$B/GW；僅供參考）", "COST_PropRev_VR_USD", "OAI_COST_PropRev_VR", "f1", ""),
             ("每 VR 等值 GW 全成本（$B/GW；僅供參考）", "COST_PropFull_VR_USD", "OAI_COST_PropFull_VR", "f1", ""),
             ("每 VR 等值 GW 差額（$B/GW；僅供參考）", "COST_PropGap_VR_USD", "OAI_COST_PropGap_VR", "f1", "")]
    tb = []
    for lab_, zr, orf, f, cls in pairs:
        for k, (who, ref) in enumerate((("智譜", zr), ("OpenAI", orf))):
            cells = []
            for y in YEARS:
                try:
                    cells.append(f"<td>{n(ref, y, f)}</td>")
                except ValueError:
                    cells.append("<td>—</td>")
            c = (cls + (" sep" if k == 0 else "")).strip()
            tb.append(f'<tr class="{c}"><td>{lab_ if k == 0 else ""}　{who}</td>{"".join(cells)}</tr>')
    S.append(f"""<section id="s9"><h2><span class="no">⑨</span>與 OpenAI v0.6 並排（美元）</h2>
<div class="must"><b>讀法提醒</b>：智譜每 VR 等值 GW 營收（2030 {n("COST_PropRev_VR_USD", Y, "f1")} $B）約為 OpenAI（{n("OAI_COST_PropRev_VR", Y, "f1")} $B）的 10 倍，經濟上不合理。招股章程更正（P1）後 η 已由約 9.5 降到 {n("CMP_Eta1H26", None, "f1")}，但供給 GW 幾乎不變（2030 {n("CMP_SupplyGW", Y, "f3")} GW），10 倍仍在——所以不能只歸因於 η：智譜的 GW 分母是「算力費 ÷ 租價」推得，而 1H26 的 token 量以 Tokenomics 物理換算所需的 GW 是其中推論部分的 {n("CMP_Eta1H26", None, "f1")} 倍（η），再乘上機隊 VR 等值係數 {n("Cost!K67", Y, "f3")}，分母可能偏小。因此<b>並排以覆蓋率與現金流為主讀，每 GW 金額僅供參考</b>，不能據此說哪一家每 GW 經濟較好。</div>
<div class="legend"><span><i style="background:var(--gap)"></i>智譜每 VR 等值 GW 差額（僅供參考）</span><span><i style="background:var(--oai)"></i>OpenAI v0.6</span></div>
<div class="cw">{chart_vs(V)}</div>
<div class="tw"><table><thead><tr><th>項目</th>{hd}</tr></thead><tbody>
{"".join(tb)}
<tr class="sep"><td>每實體 GW 差額（$B/GW）　智譜</td>{"".join(f"<td>{n('COST_PropGap_Phys_USD', y, 'f2')}</td>" for y in YEARS)}</tr>
</tbody></table></div>
<p class="note">主讀：2030 覆蓋率智譜 {n("COST_Coverage", Y, "f2")}、OpenAI {n("OAI_COST_Coverage", Y, "f2")}；規模差距很大（2030 供給 VR 等值 GW {n("CMP_Supply_VReq", Y, "f3")} 對 {n("OAI_CMP_Supply_VReq", Y, "f1")}），絕對缺口小兩個數量級。OpenAI 的問題是「錢不夠」（2030 累計外部資金需求 {n("OAI_FND_ExtNeedCum", Y, "f0")} $B）；智譜的模型基準燒錢小、已募資金多（累計 {n("FND_ExtNeedCum_USD", Y, "f2")} $B），但實際燒錢速度可能遠高於模型（第 ⑦ 節）。美元口徑＝RMB ÷ USD/CNY {n("INP_023", None, "f4")} ÷ 10；OpenAI 數字取自 OAI_Link 快照（{html.escape(str(V.get("OAI_Ref")))}），只並排、不參與計算。</p>
{V.src("COST_PropRev_VR_USD", "COST_PropFull_VR_USD", "COST_PropGap_VR_USD", "COST_Coverage", "CMP_Supply_VReq", "FND_FCF_USD", "FND_Committed_USD", "FND_ExtNeedCum_USD", "FND_CashEnd_USD", "COST_PropGap_Phys_USD", "OAI_COST_PropRev_VR", "OAI_COST_PropFull_VR", "OAI_COST_PropGap_VR", "OAI_COST_Coverage", "OAI_CMP_Supply_VReq", "OAI_FND_FCF", "OAI_FND_Committed", "OAI_FND_ExtNeedCum", "OAI_FND_CashEnd", "CMP_Eta1H26", "CMP_SupplyGW", "Cost!K67", "INP_023")}
</section>""")

    # ── ⑩ 最該審的 5 項預設 ──
    S.append(f"""<section id="s10"><h2><span class="no">⑩</span>最該審的 5 項預設</h2>
<ol>
<li><b>總算力費的估計（r2 P1）：研發項下算力費＝研發開支 × {n("INP_137", None, "pct1")}（INP_137，招股章程 1H25；FY2025、1H26 未揭露）、營業成本計算服務費＝API 銷售成本 × {n("INP_102", None, "pct0")}（INP_102）</b>：命題 1 唯一的翻轉點（≤{n("INP_137", None, "pct1", "flip_rdfee")}）；取低／高時 2030 差額 {n("COST_PropGap_VR", Y, "f0", "rdfee_lo")}／{n("COST_PropGap_VR", Y, "f0", "rdfee_hi")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "rdfee_lo")}／{n("FND_ExtNeedCum", Y, "f0", "rdfee_hi")}。</li>
<li><b>每 GW 租價下限：供應商最低毛利 {n("INP_136", None, "pct0")}（INP_136；Z5b V1）</b>：2028 起租價由下限決定；取高時 2030 差額 {n("COST_PropGap_VR", Y, "f0", "mgm_hi")}、累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "mgm_hi")}。觀察租價年變動 {n("INP_100", None, "pct0")}（INP_100）只影響 2026–2027。</li>
<li><b>η 沿用 1H26 上調值 {n("CMP_Eta1H26", None, "f2")}（INP_105＝0；η 基準 1 是 P1 的 Decision，INP_138）</b>：η 收斂至 1 時累計外部資金需求 {n("FND_ExtNeedCum", Y, "f0", "eta_conv")}——基準與該情境差距最大，也影響與 OpenAI 的每 GW 比較。</li>
<li><b>Coding Plan 額度使用率 {n("INP_060", None, "f2")}（INP_060；#13、#56、#66）</b>：擺幅大（取高／低 2030 差額 {n("COST_PropGap_VR", Y, "f0", "util_hi")}／{n("COST_PropGap_VR", Y, "f0", "util_lo")}），但主要經由 η 重校準，不是經濟機制。</li>
<li><b>2H26 API 任務數成長 +{n("INP_038", None, "pct0")}（INP_038；#15、#27）</b>：取低／高時 2030 營收淨額 {n("REV_NetCapped", Y, "f1", "task_lo")}／{n("REV_NetCapped", Y, "f1", "task_hi")}（基準 {n("REV_NetCapped", Y, "f1")}），但差額仍為負（成本同步增加）。</li>
</ol>
<p class="note">Z1–Z4 共 101 項已套用預設的全表：<code>docs/reports/20261008_v0.1.md</code> ④；r2（招股章程 P1–P5＋Z5b V1–V9）修正與前後對照：<code>docs/reports/20261008_v0.1-r2_查核修正.md</code>；敏感度原始輸出 <code>20261008_v0.1-r2_敏感度.json</code>。</p>
</section>""")

    # ── ⑪ 資料與版本 ──
    S.append(f"""<section id="s11"><h2><span class="no">⑪</span>資料與版本</h2>
<dl>
<dt>Excel</dt><dd><code>{html.escape(model.name)}</code>（v{ver}；唯一計算引擎；本頁同資料夾附 xlsx）</dd>
<dt>Tokenomics</dt><dd>{html.escape(str(tk_ver))}（<code>{html.escape(str(tk_file))}</code>），master <code>{html.escape(str(tk_sha))}</code>，讀取日 {html.escape(str(tk_date))}</dd>
<dt>OpenAI 對照</dt><dd><code>{html.escape(str(V.get("OAI_File")))}</code>（{html.escape(str(V.get("OAI_Ref")))}；SHA-256 <code>{html.escape(str(V.get("OAI_SHA256")))[:12]}…</code>）</dd>
<dt>資料日期</dt><dd>SRC_ZP 最新文件日期 {latest_src_date(model)}（含招股章程 2025-12-30 擷取 SRC_ZP_611–1054）；股價 2026-10-07；匯率 2026-09-28 中間價；模型日期 {date_disp}</dd>
<dt>檢查</dt><dd>CHK_Errors＝{n("CHK_Errors", None, "f0")}（Checks C01–C95；WARN：ARR 差距兩項、η 範圍兩項、η 上調一項）</dd>
<dt>期間與單位</dt><dd>FY2025–FY2030 曆年制，2025 為實際校準年、1H26 為第二個實際點；RMB 億；GW＝IT 關鍵電力；VR 等值＝以 VR200 Sol 層級每 GW 產能換算</dd>
</dl>
<p class="note"><span class="tag">Verified</span>已查核原文 <span class="tag">Interested-party</span>公司或利害關係方說法 <span class="tag">Analogy</span>類比推估（附區間） <span class="tag">Assumed</span>假設（附區間） <span class="tag">Derived</span>由其他數字推得 <span class="tag">Decision</span>建模決定。公司原始數據在 SRC_ZP、算力物理取自 Tokenomics（TK_Link）、假設在 Inputs、OpenAI 對照在 OAI_Link。</p>
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

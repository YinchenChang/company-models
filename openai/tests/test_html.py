"""HTML 一頁摘要（dist/）的一致性與離線測試（S6）。

1. HTML 每個帶 data-ref 的數字（含 SVG 標籤）＝ engine（pycel）以同一情境重算的 Excel 值：
   顯示文字＝同一格式化規則的結果，data-v 與 engine 值相對誤差 ≤1e-9。
2. HTML 無任何外部引用（http://、https://、//cdn、@import、url(、src=、href= 外部）。
3. dist 的 xlsx 與 model/CURRENT 位元組相同；HTML 只有一份且對應現行版本。
4. 文字敘述中的定性主張（每年差額皆為負、研發 GW 占比 6–9 成等）與 Excel 一致。
"""
from __future__ import annotations

import hashlib
import math
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tools"))
from build_html import YEARS, Locator, fmt, load_scenarios, resolve_overrides  # noqa: E402
from engine import Engine, current_model_path  # noqa: E402

MODEL = current_model_path()
STAMP = MODEL.name.split("_")[0]
VER = re.search(r"_v([\d.]+)\.xlsx$", MODEL.name).group(1).replace(".", "_")
HTML = REPO / "dist" / f"{STAMP}_OpenAI收支模型_v{VER}.html"
XLSX = REPO / "dist" / f"{STAMP}_OpenAI收支模型_v{VER}.xlsx"


class _Collect(HTMLParser):
    """收集帶 data-ref 的元素與其直接文字（略過 SVG <title> 子元素）。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.items, self.stack = [], []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if "data-ref" in a:
            self.stack.append([tag, a, [], 0])
        elif self.stack:
            if tag == "title":
                self.stack[-1][3] += 1
            self.stack[-1].append(None)

    def handle_endtag(self, tag):
        if not self.stack:
            return
        top = self.stack[-1]
        if tag == top[0] and len(top) == 4:
            self.items.append((top[1], "".join(top[2]).strip()))
            self.stack.pop()
        elif len(top) > 4:
            top.pop()
            if tag == "title":
                top[3] -= 1

    def handle_data(self, data):
        if self.stack and self.stack[-1][3] == 0:
            self.stack[-1][2].append(data)


@pytest.fixture(scope="module")
def page() -> str:
    assert HTML.exists(), f"找不到 {HTML}（請執行 python3 tools/build_html.py）"
    return HTML.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def items(page):
    p = _Collect()
    p.feed(page)
    assert p.items, "HTML 內沒有任何 data-ref 數字"
    return p.items


@pytest.fixture(scope="module")
def engine():
    return Engine(MODEL)


@pytest.fixture(scope="module")
def loc():
    return Locator(MODEL)


def _scenario_values(engine, loc, sc, keys):
    """在情境 sc 下讀 keys（(ref, year)）的 engine 值；結束後還原輸入。"""
    ov = resolve_overrides(sc, loc)
    orig = {k: engine.get(*loc.locate(k)[:2]) for k in ov}
    try:
        for k, v in ov.items():
            engine.set_key(k, v)
        out = {}
        for ref, y in keys:
            sheet, coord, _ = loc.locate(ref, y)
            out[(ref, y)] = engine.get(sheet, coord)
        return out
    finally:
        for k, v in orig.items():
            engine.set_key(k, v)


def test_html_numbers_match_excel(items, engine, loc):
    by_sc: dict[str, set] = {}
    for a, _ in items:
        y = int(a["data-y"]) if "data-y" in a else None
        by_sc.setdefault(a["data-sc"], set()).add((a["data-ref"], y))
    vals = {sc: _scenario_values(engine, loc, sc, keys) for sc, keys in by_sc.items()}
    bad, n = [], 0
    for a, text in items:
        y = int(a["data-y"]) if "data-y" in a else None
        e = vals[a["data-sc"]][(a["data-ref"], y)]
        n += 1
        want = fmt(e, a["data-fmt"])
        dv = float(a["data-v"])
        ok_v = math.isclose(dv, float(e), rel_tol=1e-9, abs_tol=1e-12)
        if text != want or not ok_v:
            bad.append((a["data-sc"], a["data-ref"], y, text, want, dv, e))
    print(f"HTML 數字比對：{n} 個（情境 {len(by_sc)} 個）")
    assert not bad, f"{len(bad)} 個不符（情境, 來源, 年, HTML, Excel 格式化, data-v, engine）：{bad[:15]}"


def test_all_scenarios_rendered(items):
    used = {a["data-sc"] for a, _ in items}
    assert set(load_scenarios()["scenarios"]) == used


def test_no_external_references(page):
    lower = page.lower()
    for pat in ("http://", "https://", "//cdn", "@import", "url(", "<script", "<link", "<img", "<iframe"):
        assert pat not in lower, f"HTML 含外部引用或外部資源標記：{pat}"
    assert not re.search(r'\b(src|href)\s*=\s*"(?!#)', page), "HTML 含 src／href 外部連結"


def test_dist_xlsx_is_current_model():
    assert XLSX.exists()
    h = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()  # noqa: E731
    assert h(XLSX) == h(MODEL)
    others = [p.name for p in (REPO / "dist").glob("*.html") if p != HTML]
    assert not others, f"dist 有非現行版本的 HTML：{others}"


def test_narrative_claims(engine):
    """HTML 文字中的定性主張。"""
    gap = engine.get_name("COST_PropGap_VR")
    assert all(g < 0 for g in gap), "『每一年差額皆為負』不成立"
    rd, dem = engine.get_name("CMP_RDGW"), engine.get_name("CMP_DemandGW")
    shares = [r / d for r, d, y in zip(rd, dem, YEARS) if y >= 2027]
    assert all(0.6 <= s <= 0.9 for s in shares), f"『2027 起研發占總需求 6–9 成』不成立：{shares}"
    assert engine.get_name("CHK_Errors") == 0

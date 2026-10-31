"""HTML 一頁摘要（dist/）的一致性與離線測試（A5；改寫自 OpenAI v0.6）。

1. HTML 每個帶 data-ref 的數字（含 SVG 標籤）＝ engine（pycel）以同一情境重算的 Excel 值：
   顯示文字＝同一格式化規則的結果，data-v 與 engine 值相對誤差 ≤1e-9。
2. HTML 無任何外部引用（http://、https://、//cdn、@import、url(、src=、href= 外部）。
3. dist 的 xlsx 與 model/CURRENT 位元組相同；HTML 只有一份且對應現行版本。
4. 文字敘述中的定性主張（2026–28 差額為負、2029 起轉正、現金谷底在 2028 等）與 Excel 一致。
"""
from __future__ import annotations

import ast
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
from engine import Engine, current_model_path, resolve_key  # noqa: E402

MODEL = current_model_path()
STAMP = MODEL.name.split("_")[0]
VER = re.search(r"_v([\d.]+)\.xlsx$", MODEL.name).group(1).replace(".", "_")
HTML = REPO / "dist" / f"{STAMP}_Anthropic收支模型_v{VER}.html"
XLSX = REPO / "dist" / f"{STAMP}_Anthropic收支模型_v{VER}.xlsx"


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
    orig = {k: engine.get(*resolve_key(loc.names, k)) for k in ov}   # v0.1.2：覆寫鍵可含 NAME[k]（TK 快照多格名稱）
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
        dv = ast.literal_eval(a["data-v"])
        if isinstance(e, str) or isinstance(dv, str):
            ok_v = dv == e
        else:
            ok_v = math.isclose(float(dv), float(e), rel_tol=1e-9, abs_tol=1e-12)
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
    """HTML 文字中的定性主張（①②④⑤⑥⑦；v0.1.7 起：2026 為正、2027–2029 為負、2030 轉正；累計外部資金需求 0；現金谷底在 2029）。"""
    gap = dict(zip(YEARS, engine.get_name("COST_PropGap_VR")))
    assert gap[2026] > 0, "『2026 年可覆蓋全成本』不成立"
    assert all(gap[y] < 0 for y in (2027, 2028, 2029)), "『2027–2029 差額為負』不成立"
    assert gap[2030] > 0, "『2030 年轉正』不成立"
    for nm in ("COST_PropGapEcon_VR", "COST_PropGap_GW"):
        g = dict(zip(YEARS, engine.get_name(nm)))
        assert g[2026] > 0 and g[2029] < 0 < g[2030], f"『三種口徑正負一致』不成立：{nm}"
    assert all(math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9) for a, b in zip(engine.get_name("COST_PropGapEcon_VR"), engine.get_name("COST_PropGap_VR"))), "『自建 0 → 現金與經濟口徑相同』不成立"
    assert all(v in (None, "", 0) for v in engine.get_name("FND_ExtNeedCum")), "『累計外部資金需求 0』不成立"
    fcf = dict(zip(YEARS, engine.get_name("FND_FCF")))
    assert fcf[2026] > 0 and all(fcf[y] < 0 for y in (2027, 2028, 2029)) and fcf[2030] > 0, "『營運自由現金流 2026 正、2027–2029 負、2030 正』不成立"
    cash = dict(zip(YEARS, engine.get_name("FND_CashNoExt")))
    assert cash[2029] == engine.get_name("FND_CashNoExtTrough")
    sup = dict(zip(YEARS, engine.get_name("CMP_SupplyGW")))
    assert sup[2030] < sup[2028], "『供給 2028 年後遞減』不成立"
    assert all(g < 0 for g in engine.get_name("OAI_COST_PropGap_VR")), "『OpenAI 每年差額皆為負』不成立"
    assert engine.get_name("COST_S1fNonCompDiff") > engine.get_name("INP_214"), "『W4 推得值超過 30% 門檻』不成立"
    assert engine.get_name("CHK_Errors") == 0


def test_single_driver_claims(engine, loc):
    """⑥ 的單一驅動清單由 Excel 值產生（HTML 只取值）；v0.1.7 起至少一個單一驅動會產生外部資金需求、組合 combo1 也會；歷史情境 v016 不會。"""
    single = [k for k, g in load_scenarios()["scenarios"].items() if len(g) == 1]
    need = {sc: _scenario_values(engine, loc, sc, [("FND_ExtNeedCum", 2030)])[("FND_ExtNeedCum", 2030)] for sc in single}
    assert any(v and v > 0 for v in need.values())
    for sc in ("api_task_lo", "elast_weak", "price_hi", "s1u_fast"):          # ① 備註所列
        assert need[sc] > 0, sc
    for sc in ("combo1", "v016"):
        v = _scenario_values(engine, loc, sc, [("FND_ExtNeedCum", 2030)])[("FND_ExtNeedCum", 2030)]
        assert (v > 0) if sc == "combo1" else (v in (None, "", 0)), (sc, v)

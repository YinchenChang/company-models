"""builder 與資料分層的結構性測試（不需 LibreOffice）：E6、穩定 ID、OAI_Link 隔離、SRC／Inputs 與 yaml 一致、Derived 列公式化。"""
import json
import re
import sys
from pathlib import Path

import openpyxl
import pytest
import yaml

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "builder"))
from build import DERIVED, CALC_SHEETS  # noqa: E402
from common import e6_violations  # noqa: E402
from engine import Engine, current_model_path  # noqa: E402

SRC = yaml.safe_load((REPO / "data" / "anthropic_src.yaml").read_text(encoding="utf-8"))["rows"]
INP = yaml.safe_load((REPO / "data" / "anthropic_inputs.yaml").read_text(encoding="utf-8"))["rows"]
REG = json.loads((REPO / "builder" / "id_registry.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def wb():
    return openpyxl.load_workbook(current_model_path())


@pytest.fixture(scope="module")
def eng():
    return Engine(current_model_path())


def _same(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-12 * max(abs(a), abs(b), 1)
    return a == b


def test_e6_no_constants_in_calc_sheets(wb):
    """E6：計算頁（A2：Demand、Revenue；A3：Compute、Cost）公式不得內含常數；恆等式只容許 1−比例、1＋成長率、年數 +1；定義常數在 Inputs。"""
    present = [s for s in CALC_SHEETS if s in wb.sheetnames]
    assert {"Demand", "Revenue", "Compute", "Cost", "Funding", "Reverse"} <= set(present)
    bad = [c for s in present for c in e6_violations(wb[s])]
    assert bad == [], bad[:10]
    n = sum(1 for s in present for row in wb[s].iter_rows(min_row=5) for c in row if isinstance(c.value, str) and c.value.startswith("="))
    assert n >= 800


def test_oai_link_not_referenced_by_calc_sheets(wb):
    """D20：OAI_Link 只被 Checks 引用。"""
    for ws in wb.worksheets:
        if ws.title in ("OAI_Link", "Checks"):
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert "OAI_" not in c.value, f"{ws.title}!{c.coordinate} 引用 OAI_Link：{c.value}"


def test_stable_ids():
    """SRC_ANT_nnn 與 yaml 一致且不重複；INP_nnn、Cnn 由 registry 發出、不重用退役號碼。"""
    ids = [r["id"] for r in SRC]
    assert len(ids) == len(set(ids))
    assert ids == sorted(ids) and ids[-1] == f"SRC_ANT_{len(ids):03d}"
    assert all(REG["SRC"][r["key"]] == r["id"] for r in SRC)
    for kind in ("INP", "CHK"):
        vals = list(REG[kind].values())
        assert len(vals) == len(set(vals))
        assert not set(vals) & set(REG.get(f"retired_{kind}", []))
    assert {r["key"] for r in INP} == set(REG["INP"])
    assert set(REG.get("retired_INP", [])) >= {"INP_016", "INP_067", "INP_068", "INP_069"}     # A2：每任務 token 改取 Tokenomics；A3：容量上限佔位改接 Compute


def test_src_and_inputs_values_equal_yaml(eng):
    """SRC_ANT／Inputs 每個有值的列：Excel 具名範圍的值等於 yaml（Derived 公式列除外，另測）。"""
    for r in SRC:
        if r["value"] is not None and not DERIVED.get(r["key"]):
            assert _same(eng.get_name(r["id"]), r["value"]), r["id"]
    for r in INP:
        assert _same(eng.get_name(REG["INP"][r["key"]]), r["value"]), r["key"]


def test_derived_rows_are_formulas(wb):
    """yaml 檔頭：標記 Derived 的 SRC 列改為公式（或保留報導值並註明）；公式只引用 SRC_ANT／INP 具名範圍。"""
    ws = wb["SRC_ANT"]
    rows = {ws.cell(r, 2).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 2).value}
    for r in SRC:
        if r["tag"] != "Derived":
            continue
        n = rows[r["key"]]
        f = ws[f"E{n}"].value
        if DERIVED[r["key"]]:
            assert isinstance(f, str) and f.startswith("=") and re.fullmatch(r"[=+\-*/()A-Z_0-9.1]+", f.replace("SRC_ANT_", "S").replace("INP_", "I")), (r["id"], f)
            assert ws[f"V{n}"].value == "Derived→公式"
        else:
            assert ws[f"V{n}"].value.startswith("Derived→保留報導值")


def test_inputs_ranges_and_tags():
    """共同規則第 4 節：Analogy／Assumed 一律給區間；低 ≤ 值 ≤ 高；標記只限 Assumed／Analogy／Decision。"""
    for r in INP:
        assert r["tag"] in ("Assumed", "Analogy", "Decision"), r["key"]
        assert r["lo"] is not None and r["hi"] is not None, r["key"]
        if isinstance(r["value"], str):          # 文字鍵（Tokenomics 任務名稱）：Decision，低＝高＝值
            assert r["tag"] == "Decision" and r["lo"] == r["hi"] == r["value"], r["key"]
        else:
            assert r["lo"] <= r["value"] <= r["hi"], r["key"]
        assert r.get("basis"), r["key"]


def test_named_outputs_for_a3(eng):
    """A3／A4 會用到的具名範圍（與 OpenAI v0.6 同名者）存在且為 6 年。"""
    for n in ("DEM_Tok_Paid_Top", "DEM_Tok_Paid_Mid", "DEM_Tok_Paid_Low", "DEM_Tok_Free_Top", "DEM_Tok_Free_Mid", "DEM_Tok_Free_Low",
              "DEM_Tok_Paid", "DEM_Tok_Free", "DEM_Tok_Total", "DEM_Tok_API", "REV_Sub", "REV_API", "REV_Gross", "REV_PartnerShare", "REV_Net",
              "REV_GrossCapped", "REV_NetCapped", "REV_CapFactor", "REV_Consumer", "REV_Enterprise"):
        v = eng.get_name(n)
        assert isinstance(v, list) and len(v) == 6 and all(isinstance(x, (int, float)) for x in v), n
    assert eng.get_name("CHK_Errors") == 0


def test_named_outputs_for_a4(eng):
    """A3 輸出、A4 會用到的具名範圍（與 OpenAI v0.6 同名）存在且為 6 年；容量上限 2025、2026＝1；Revenue 與 Compute 一致。"""
    for n in ("CMP_SupplyGW", "CMP_Supply_VReq", "CMP_InfGW_Eff", "CMP_InfGW", "CMP_RDGW", "CMP_DemandGW", "CMP_InfAvailGW", "CMP_CapFactor", "CMP_Eta",
              "COST_Compute", "COST_ContractPay", "COST_OwnedCapex", "COST_NonCompExSBC", "COST_SBC", "COST_FullCash", "COST_PropRev_VR",
              "COST_PropFull_VR", "COST_PropGap_VR", "COST_Coverage", "COST_CloudGM", "COST_ComputeEcon", "COST_GapCash"):
        v = eng.get_name(n)
        assert isinstance(v, list) and len(v) == 6 and all(isinstance(x, (int, float)) for x in v), n
    cap = eng.get_name("CMP_CapFactor")
    assert cap[0] == 1 and cap[1] == 1
    assert eng.get_name("REV_CapFactor") == cap
    assert abs(eng.get_name("COST_Gap2025")) < 1e-9                     # 2025 算力成本＝說明書（v0.1.7 起正式 S-1 7.326）


def test_reverse_not_fed_back(wb):
    """D19：Reverse（管理層目標反解）只作對照，不回饋基準——除 Reverse 本頁與 Checks 外，任何公式不得引用 Reverse! 或 RVS_ 名稱。"""
    assert "Reverse" in wb.sheetnames
    for ws in wb.worksheets:
        if ws.title in ("Reverse", "Checks"):
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert "Reverse!" not in c.value and "RVS_" not in c.value, f"{ws.title}!{c.coordinate} 引用 Reverse：{c.value}"


def test_named_outputs_v01(eng):
    """v0.1（A4）命題 2 與反向模式輸出（與 OpenAI v0.6 同名）存在；2025 為實際年（流量欄空白），2026–2030 為數值；外部資金需求非負、年底現金 ≥ 最低現金。"""
    for n in ("FND_FCF", "FND_Committed", "FND_EquityIn", "FND_MinCash", "FND_CashEnd", "RVS_Target", "RVS_MultAPI", "RVS_MultSub", "RVS_MultProp", "RVS_FCF",
              "COST_PropGap_GW", "COST_ComputeAnchor", "COST_PropGap_VR_Anchor", "COST_LeaseRent", "COST_PropGap_VR_Rent", "COST_PropGap_VR_D22"):
        v = eng.get_name(n)
        assert isinstance(v, list) and len(v) == 6 and all(isinstance(x, (int, float)) for x in v), n
    for n in ("FND_ExtNeed", "FND_ExtNeedCum", "FND_CashOpen", "FND_Headroom", "RVS_ExtNeedCum", "FND_ExtNeedCumIPO", "FND_ExtNeedCumCond", "FND_ExtNeedCumAll",
              "FND_ExtNeedCumAnchor", "FND_ExtNeedCumRent", "FND_ExtNeedCumD22", "FND_ExtNeedCumAdverse"):
        v = eng.get_name(n)
        assert len(v) == 6 and all(isinstance(x, (int, float)) for x in v[1:]), n
    ext, end, mn = eng.get_name("FND_ExtNeed"), eng.get_name("FND_CashEnd"), eng.get_name("FND_MinCash")
    assert all(x >= 0 for x in ext[1:])
    assert all(e >= m - 1e-9 for e, m in zip(end[1:], mn[1:]))
    for n in ("FND_FirstGapYear", "FND_ExtNeedPeak", "FND_PeakYear", "FND_Contingent", "FND_CommitAfter2030", "COST_PartnerGapMax"):
        assert eng.get_name(n) not in (None, ""), n
    assert eng.get_name("CHK_Errors") == 0


# ── v0.1.3 B：Bridge、EtaDiag（唯讀診斷頁）─────────────────────────────────
def test_bridge_not_fed_back(wb):
    """Bridge、EtaDiag 只作對照與診斷：除本頁與 Checks 外，任何公式不得引用 Bridge!、EtaDiag! 或 ETA_ 名稱。"""
    assert "Bridge" in wb.sheetnames and "EtaDiag" in wb.sheetnames
    assert wb.sheetnames.index("Bridge") == wb.sheetnames.index("Cost") + 1       # 放在 Cost 之後
    for ws in wb.worksheets:
        if ws.title in ("Bridge", "EtaDiag", "Checks"):
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert not re.search(r"\b(Bridge|EtaDiag)!|\bETA_", c.value), f"{ws.title}!{c.coordinate} 引用診斷頁：{c.value}"


def test_eta_decomposition_does_not_change_eta(eng):
    """η 拆解只作診斷：η 等於 Compute 的值（v0.1.9 起為 0.603958469653879：TK v21.10 的 TPU v7、Trainium3 產出比改 FP8 對齊（0.8 → 0.66、0.5 → 0.28），2025 token 換算 GW 上升；
    v0.1.7–v0.1.8 為 0.46763641507486053：W3 2025 校準改正式 S-1、W5 TK v21.7 的 TPU／Trainium 持有比下降使既有雲端加權合約價下降；
    v0.1.4–v0.1.6 為 0.525427732668394；v0.1.3 為 0.52729641508883，差異全來自 TK IF_TokGW_Luna）；乘積 × 殘差 ＝ η；殘差為正。"""
    eta = eng.get_name("CMP_Eta2025")
    assert abs(eta - 0.603958469653879) < 1e-9
    assert abs(eng.get_name("ETA_Eta") - eta) < 1e-12
    prod = eng.get_name("ETA_fP") * eng.get_name("ETA_fU") * eng.get_name("ETA_fW")
    assert abs(prod - eng.get_name("ETA_Prod")) < 1e-12
    assert abs(prod * eng.get_name("ETA_Resid") - eta) < 1e-12
    assert eng.get_name("ETA_Resid") > 0
    assert 0 < eng.get_name("ETA_fU") < 2 and eng.get_name("ETA_fW") > 0 and eng.get_name("ETA_fP") > 0


def test_eta_prod_derate_section(eng):
    """v0.1.4：f_U'＝f_U × TK 生產折減；殘差' ＝ η ÷ (f_P × f_U' × f_W)；η 不變，殘差 ×（1 ÷ 生產折減）。"""
    eta = eng.get_name("CMP_Eta2025")
    pd = eng.get_name("TK_IF_ProdDerate")
    assert abs(eng.get_name("ETA_fU_Prod") - eng.get_name("ETA_fU") * pd) < 1e-12
    prod2 = eng.get_name("ETA_fP") * eng.get_name("ETA_fU_Prod") * eng.get_name("ETA_fW")
    assert abs(prod2 - eng.get_name("ETA_Prod2")) < 1e-12
    assert abs(prod2 * eng.get_name("ETA_Resid2") - eta) < 1e-12
    assert abs(eng.get_name("ETA_Resid2") - eng.get_name("ETA_Resid") / pd) < 1e-12


def test_bridge_ties_to_calc_sheets(eng):
    """Bridge 逐層列的值與計算頁一致（核對列為 0）；有 Bridge 時 Checks e1_* 全 OK。"""
    rows = eng.get("Bridge", "A8:I140")
    byid = {r[0]: r for r in rows if r[0] != ""}
    for rid, r in byid.items():
        lab = str(r[1])
        if lab.startswith("核對"):
            assert all(abs(x) < 1e-9 for x in r[3:9] if isinstance(x, (int, float))), f"{rid} {lab}"
    gap = eng.get_name("COST_PropGap_VR")
    gap_row = next(r for r in byid.values() if str(r[1]).startswith("每 VR 等值 GW：差額（含股權報酬；命題 1）"))
    assert all(abs(a - b) < 1e-9 for a, b in zip(gap_row[3:9], gap))
    sup_row = next(r for r in byid.values() if r[1] == "供給 GW 合計")
    assert all(abs(a - b) < 1e-9 for a, b in zip(sup_row[3:9], eng.get_name("CMP_SupplyGW")))


def _s1f_rows(wb):
    """Cost 第二十四節（v0.1.6 正式 S-1 對照）的列號範圍：節標題至下一節（v0.1.7 第二十五節）之前；無下一節時至頁尾。"""
    ws = wb["Cost"]
    start = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("二十四、"))
    end = next((r for r in range(start + 1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("二十五、")), ws.max_row + 1)
    return ws, set(range(start, end))


def test_s1f_section_not_fed_back(wb):
    """v0.1.6 W3：Cost 第二十四節只對照（不回饋基準）——除本節與 Checks 外，任何公式不得引用本節儲存格或本節具名範圍。"""
    ws, rows = _s1f_rows(wb)
    names = [n for n, v in wb.defined_names.items() if v.attr_text.startswith("Cost!")
             and int(re.search(r"\$(\d+)", v.attr_text).group(1)) in rows]
    assert "COST_Eta2026H1" in names and "COST_PropGap_VR_EtaH1" in names
    for sh in wb.worksheets:
        if sh.title == "Checks":
            continue
        for row in sh.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")) or (sh.title == "Cost" and c.row in rows):
                    continue
                refs = [int(x) for x in re.findall(r"Cost!\$?[A-Z]{1,3}\$?(\d+)", v)]
                if sh.title == "Cost":
                    refs += [int(x) for x in re.findall(r"(?<![A-Za-z!_$])\$?[D-I]\$?(\d+)\b", re.sub(r"[A-Za-z_]+![$A-Z0-9:]+", "", v))]
                assert not [r for r in refs if r in rows], f"{sh.title}!{c.coordinate} 引用第二十四節：{v}"
                assert not [n for n in names if re.search(rf"\b{n}\b", v)], f"{sh.title}!{c.coordinate} 引用第二十四節具名範圍：{v}"


def test_s1f_w2_and_eta_unchanged(eng):
    """v0.1.6 W2：2026 營收基準＝正式 S-1 上半年＋Q3＋Q4（Q3 持平）；基準 η 不因 v0.1.6 W3 改變（v0.1.7 起值見 test_eta_decomposition_does_not_change_eta）；B 情境 2025 欄＝基準。"""
    s = {r["key"]: r["id"] for r in SRC}
    i = {r["key"]: REG["INP"][r["key"]] for r in INP}
    assert eng.get_name(i["s1f_switch"]) == 1 and eng.get_name(i["u1_switch"]) == 1
    q3 = eng.get_name(s["rev_s1a_2026q3"]) if eng.get_name(i["s1a_q3_switch"]) == 1 else eng.get_name(s["rev_2026_q3"])     # v0.1.8 R5a：Q3 改 S-1/A 實績
    tgt = eng.get_name(s["rev_s1f_2026h1"]) + q3 * (2 + eng.get_name(i["q4_growth"]))
    assert abs(eng.get_name("REV_Target")[1] - tgt) < 1e-9
    assert abs(eng.get_name("CMP_Eta2025") - 0.603958469653879) < 1e-9          # v0.1.9：TK v21.10（v0.1.7–v0.1.8 為 0.46763641507486053）
    assert eng.get_name("COST_Eta2026H1") > 0
    for a, b in ((eng.get_name("COST_PropGap_VR_EtaH1")[0], eng.get_name("COST_PropGap_VR")[0]),
                 (eng.get_name("COST_Coverage_EtaH1")[0], eng.get_name("COST_Coverage")[0])):
        assert abs(a - b) < 1e-9


# ═════════════ v0.1.7（工作單 20261031 v0.1.7） ═════════════
def _ids():
    return {r["key"]: r["id"] for r in SRC}, {r["key"]: REG["INP"][r["key"]] for r in INP}


def test_v017_w1_unallocated_balance(eng):
    """W1(b)：未歸屬餘額＝S-1 三段合計 − 模型逐家合約總額（不含 AMD）；供給 GW＝付款 ÷ VR200 合約價。
    v0.1.8：三段改 S-1/A 9-30 口徑（R5b：表內＋未起租＋9-30 後），付款＝餘額 × S-1/A 結構合成年度比例（R5c）；附條件扣除額基準 0（R5d）。"""
    s, i = _ids()
    assert eng.get_name(i["s1a_commit_switch"]) == 1 and eng.get_name(i["s1a_unalloc_switch"]) == 1 and eng.get_name(i["s1a_cond_deduct_switch"]) == 0
    seg = sum(eng.get_name(s[k]) for k in ("cmt_s1a_table", "cmt_s1a_leases_nc", "cmt_s1a_post0930"))
    assert abs(seg - eng.get_name("COST_S1fCommit3")) < 1e-9 and abs(seg - eng.get_name("COST_S1aCommit3")) < 1e-9
    bal = eng.get_name("COST_S1Unalloc")
    assert abs(bal - max(0.0, seg - eng.get_name("COST_S1fCommitModel"))) < 1e-9 and bal > 0
    pay, sh = eng.get_name("COST_Pay_S1Unalloc"), eng.get_name("COST_S1aUnallocShare")
    assert all(abs(v - bal * w) < 1e-9 for v, w in zip(pay, sh))
    price = eng.get_name("CMP_Price_VR200")
    assert all(abs(g - p / pp) < 1e-9 for g, p, pp in zip(eng.get_name("CMP_Supply_S1Unalloc"), pay, price))
    assert eng.get_name("COST_GoogleAddl") == 0 and eng.get_name("COST_S1fSpaceXFactor") == 1     # 工作單預設：Google 額外 TPU 不進基準；SpaceXAI 沿用月費


def test_v017_w2_own_capex_gw_method(eng):
    """W2：自建 GW（Wattonomics v0.1.7：各年 0）× TK 每 GW 建置成本 → 自建資本支出與自建 GW 為 0；Fluidstack $50B 不另計。"""
    _s, i = _ids()
    assert eng.get_name(i["own_capex_method"]) == 1
    assert all(eng.get_name(i[f"own_gw_add_{y}"]) == 0 for y in range(2025, 2031))
    assert all(v == 0 for v in eng.get_name("COST_OwnedCapex")) and all(v == 0 for v in eng.get_name("CMP_Supply_Owned"))


def test_v017_w4_noncompute_calibration(eng):
    """W4（核對後）：2026 非算力成本（不含股權報酬）＝上半年（營收 − 算力 − 營業損益 − 股權報酬代理）＋下半年 Q3 × 2 − 平台抽成；股權報酬＝S-1 調整數代理。
    v0.1.8 R5a：Q3＝S-1/A 實際（營收 − 算力 − 營業損益 − 股權報酬 − 非現金捐贈）；股權報酬代理下半年＝Q3 實際（股權報酬＋非現金捐贈）× 2。"""
    s, i = _ids()
    g = lambda k: eng.get_name(s[k])  # noqa: E731
    assert eng.get_name(i["s1f_d22_q3_switch"]) == 1 and eng.get_name(i["s1f_sbc_proxy_switch"]) == 1 and eng.get_name(i["s1a_q3_switch"]) == 1
    sbc_h1 = g("adjoi_s1f_2026h1") - g("oi_s1f_2026h1")
    sbc_q3 = g("sbc_s1a_2026q3") + g("charity_s1a_2026q3")
    assert abs(eng.get_name("COST_S1fSBCProxy2026") - (sbc_h1 + 2 * sbc_q3)) < 1e-9
    assert abs(eng.get_name("COST_SBC")[1] - (sbc_h1 + 2 * sbc_q3)) < 1e-9
    q3 = g("rev_s1a_2026q3") - g("ci_s1a_2026q3") - g("oi_s1a_2026q3") - g("sbc_s1a_2026q3") - g("charity_s1a_2026q3")
    assert abs(eng.get_name("COST_S1aQ3NonComp") - q3) < 1e-9
    h1 = g("rev_s1f_2026h1") - g("ci_s1f_2026h1") - g("oi_s1f_2026h1") - sbc_h1
    want = h1 + 2 * q3 - eng.get_name("REV_PartnerShare")[1]
    assert abs(eng.get_name("COST_S1fNonComp2026") - want) < 1e-9
    assert abs(eng.get_name("COST_NonCompExSBC")[1] - want) < 1e-9
    assert eng.get_name("COST_S1fAMDFactor") == 1                       # 基準：AMD 維持模型 GW 型
    assert abs(eng.get_name("REV_Target")[1] - (g("rev_s1f_2026h1") + 2 * g("rev_s1a_2026q3"))) < 1e-9     # R5a：營收基準 Q3 實績、Q4＝Q3


def test_v017_switches_off_reproduce_v016(eng):
    """各項開關設 0（含 v0.1.7.1 的 V2、V3 開關）、TK 快照回到 v21.2（值取自 model/archive 的 v0.1.6）時，命題輸出與 v0.1.6 逐位相同（回歸）。"""
    old = REPO / "model" / "archive" / "20261031_Anthropic_v0.1.6.xlsx"
    wo = openpyxl.load_workbook(old)
    wv = openpyxl.load_workbook(old, data_only=True)
    _s, i = _ids()
    ov = {i[k]: 0 for k in ("s1f_commit_switch", "own_capex_method", "s1f_cal25_switch", "s1f_d22_switch", "s1f_d22_q3_switch", "s1f_sbc_proxy_switch",
                            "googl_upfront_switch", "s1f_round_switch", "s1f_cond_switch",          # v0.1.7.1：V2、V3 開關一併關閉
                            "s1a_q3_switch", "s1a_commit_switch", "s1a_unalloc_switch")}             # v0.1.8：R5a–R5c 開關一併關閉（R4、R5d 基準已為 0）
    for n, d in wo.defined_names.items():
        if not n.startswith("TK_") or n in ("TK_File", "TK_Version", "TK_Commit", "TK_ReadDate", "TK_HdrGen", "TK_HdrCost") or n not in eng.names:
            continue
        sh, ref = d.attr_text.split("!")
        cells = wo[sh][ref.replace("$", "")]
        if isinstance(cells, tuple):
            ov.update({f"{n}[{k + 1}]": c.value for k, c in enumerate(x for row in cells for x in row)})
        else:
            ov[n] = cells.value
    from engine import resolve_key
    orig = {k: eng.get(*resolve_key(eng.names, k)) for k in ov}
    try:
        for k, v in ov.items():
            eng.set_key(k, v)
        for n in ("COST_PropGap_VR", "COST_Coverage", "FND_ExtNeedCum", "FND_CashNoExt", "CMP_Eta2025", "COST_Compute", "CMP_SupplyGW", "REV_Gross", "COST_NonCompExSBC"):
            sh, ref = wo.defined_names[n].attr_text.split("!")
            cells = wv[sh][ref.replace("$", "")]
            want = [c.value for row in cells for c in row] if isinstance(cells, tuple) else [cells.value]
            got = eng.get_name(n)
            got = got if isinstance(got, list) else [got]
            for a, b in zip(got, want):
                if isinstance(b, (int, float)):
                    assert abs(a - b) <= 1e-9 * max(1, abs(b)), (n, a, b)
                else:
                    assert a in (None, "") and b in (None, ""), (n, a, b)
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)


def test_v0171_v2_committed_matches_s1(eng):
    """v0.1.7.1 V2：2026 已到位融資＝Series G 30.2＋Amazon G-2 5＋Series H 65.2＋Google G-2（38.6 百萬股 × 259.1364 ÷ 1000），對 S-1 特別股發行收入 110.432 差 ≤ 0.5。"""
    s, i = _ids()
    g = lambda k: eng.get_name(s[k])  # noqa: E731
    assert eng.get_name(i["googl_upfront_switch"]) == 1 and eng.get_name(i["s1f_round_switch"]) == 1
    want = g("fund_s1f_series_g_q1") + g("strat_amzn_2026_series_g") + g("fund_s1f_series_h_q2") + g("fund_s1f_g2_googl_shares") * g("fund_s1f_g_price") / eng.get_name(i["m_to_b"])
    com = eng.get_name("FND_Committed")
    assert abs(com[1] - want) < 1e-9 and all(v == 0 for k, v in enumerate(com) if k != 1)
    assert abs(com[1] - g("fund_s1f_pref_2026h1")) <= eng.get_name(i["s1f_cmt_tol"])
    assert abs(eng.get_name("FND_CommittedS1Gap") - (com[1] - g("fund_s1f_pref_2026h1"))) < 1e-9


def test_v0171_switches_off_reproduce_v017(eng):
    """v0.1.7.1 的 V2、V3 開關設 0 時，v0.1.7 的全部計算具名範圍（DEM_／REV_／CMP_／COST_／FND_／RVS_）與 v0.1.7 逐位相同（範圍外變動 0 的回歸）。
    V1（SRC_ANT_427–431 原文精度）只影響對照列（Cost 第二十五節 W4 交叉核對、Checks），不在這些名稱內；Checks 另由報告逐格比對。"""
    old = REPO / "model" / "archive" / "20261031_Anthropic_v0.1.7.xlsx"
    wo = openpyxl.load_workbook(old)
    wv = openpyxl.load_workbook(old, data_only=True)
    _s, i = _ids()
    ov = {i[k]: 0 for k in ("googl_upfront_switch", "s1f_round_switch", "s1f_cond_switch", "s1a_q3_switch", "s1a_commit_switch", "s1a_unalloc_switch")}  # v0.1.8 開關一併關閉
    ov.update(_tk_ov(eng, V018_XLSX))   # v0.1.9：TK 快照回到 v21.7（v0.1.7 當時；值取自 v0.1.8 Excel）
    from engine import resolve_key
    orig = {k: eng.get(*resolve_key(eng.names, k)) for k in ov}
    n_cmp = 0
    try:
        for k, v in ov.items():
            eng.set_key(k, v)
        for n, d in wo.defined_names.items():
            if not n.startswith(("DEM_", "REV_", "CMP_", "COST_", "FND_", "RVS_")):
                continue
            sh, ref = d.attr_text.split("!")
            cells = wv[sh][ref.replace("$", "")]
            want = [c.value for row in cells for c in row] if isinstance(cells, tuple) else [cells.value]
            got = eng.get_name(n)
            got = got if isinstance(got, list) else [got]
            assert len(got) == len(want), n
            for a, b in zip(got, want):
                n_cmp += 1
                if isinstance(b, (int, float)) and not isinstance(b, bool):
                    assert abs(a - b) <= 1e-9 * max(1, abs(b)), (n, a, b)
                else:
                    assert (a in (None, "") and b in (None, "")) or a == b, (n, a, b)
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)
    assert n_cmp > 1500          # v0.1.7.1：1,696 個值


V018_XLSX = REPO / "model" / "archive" / "20261103_Anthropic_v0.1.8.xlsx"
TK_META = ("TK_File", "TK_Version", "TK_Commit", "TK_ReadDate", "TK_HdrGen", "TK_HdrCost")


def _tk_ov(eng, old):
    """old 活頁簿 TK_Link 的全部具名值（覆寫用；v0.1.9 起歷史回歸需把 TK 快照一併回到當時版本）。"""
    wo = openpyxl.load_workbook(old)
    ov = {}
    for n, d in wo.defined_names.items():
        if not n.startswith("TK_") or n in TK_META or n not in eng.names:
            continue
        sh, ref = d.attr_text.split("!")
        cells = wo[sh][ref.replace("$", "")]
        if isinstance(cells, tuple):
            ov.update({f"{n}[{k + 1}]": c.value for k, c in enumerate(x for row in cells for x in row)})
        else:
            ov[n] = cells.value
    return ov


def _reproduce(eng, old, ov, prefixes=("DEM_", "REV_", "CMP_", "COST_", "FND_", "RVS_")):
    wo = openpyxl.load_workbook(old)
    wv = openpyxl.load_workbook(old, data_only=True)
    from engine import resolve_key
    orig = {k: eng.get(*resolve_key(eng.names, k)) for k in ov}
    n_cmp = 0
    try:
        for k, v in ov.items():
            eng.set_key(k, v)
        for n, d in wo.defined_names.items():
            if not n.startswith(prefixes):
                continue
            sh, ref = d.attr_text.split("!")
            cells = wv[sh][ref.replace("$", "")]
            want = [c.value for row in cells for c in row] if isinstance(cells, tuple) else [cells.value]
            got = eng.get_name(n)
            got = got if isinstance(got, list) else [got]
            assert len(got) == len(want), n
            for a, b in zip(got, want):
                n_cmp += 1
                if isinstance(b, (int, float)) and not isinstance(b, bool):
                    assert abs(a - b) <= 1e-9 * max(1, abs(b)), (n, a, b)
                else:
                    assert (a in (None, "") and b in (None, "")) or a == b, (n, a, b)
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)
    return n_cmp


def test_v018_switches_off_reproduce_v0171(eng):
    """v0.1.8 的 R5a、R5b、R5c 開關設 0（R4、R5d 基準已為 0）時，v0.1.7.1 的全部計算具名範圍（DEM_／REV_／CMP_／COST_／FND_／RVS_）逐位相同（範圍外變動 0 的回歸）。"""
    _s, i = _ids()
    assert eng.get_name(i["s1f_spacex_fee_switch"]) == 0 and eng.get_name(i["s1a_cond_deduct_switch"]) == 0
    ov = {i[k]: 0 for k in ("s1a_q3_switch", "s1a_commit_switch", "s1a_unalloc_switch")}
    ov.update(_tk_ov(eng, V018_XLSX))   # v0.1.9：TK 快照回到 v21.7（v0.1.7.1 當時）
    n_cmp = _reproduce(eng, REPO / "model" / "archive" / "20261102_Anthropic_v0.1.7.1.xlsx", ov)
    assert n_cmp > 1700          # v0.1.7.1 的計算具名範圍值


def test_v018_r5c_shares_and_r4_path(eng):
    """R5c：年度比例＝（表內逐年＋未起租與 9-30 後〔起租年起、加權租期〕）÷ 三段合計，全期合計 1；R4：S-1 路徑合計＝上限 120.5，開關 1 時進 SpaceX 付款與合約總額。"""
    s, i = _ids()
    g = lambda k: eng.get_name(s[k])  # noqa: E731
    seg = g("cmt_s1a_table") + g("cmt_s1a_leases_nc") + g("cmt_s1a_post0930")
    tb = [0.0] + [g(k) for k in ("cmt_s1a_tbl_2026q4", "cmt_s1a_tbl_2027", "cmt_s1a_tbl_2028", "cmt_s1a_tbl_2029", "cmt_s1a_tbl_2030")]
    ls, wt = eng.get_name(i["s1a_lease_start"]), g("cmt_s1a_lease_wterm")
    lpy = (g("cmt_s1a_leases_nc") + g("cmt_s1a_post0930")) / wt
    for y, t, w in zip(range(2025, 2031), tb, eng.get_name("COST_S1aUnallocShare")):
        assert abs(w - (t + (lpy if ls <= y < ls + wt else 0)) / seg) < 1e-12
    assert abs((sum(tb) + g("cmt_s1a_tbl_after") + lpy * wt) / seg - 1) < 1e-9
    path = eng.get_name("COST_S1aSpaceXPath")
    assert abs(sum(path) - g("cmt_s1f_spacex")) < 1e-9 and abs(eng.get_name("COST_S1aSpaceXTot") - g("cmt_s1f_spacex")) < 1e-9
    from engine import resolve_key
    k = i["s1f_spacex_fee_switch"]
    orig = eng.get(*resolve_key(eng.names, k))
    try:
        eng.set_key(k, 1)
        assert all(abs(a - b) < 1e-9 for a, b in zip(eng.get_name("COST_Pay_Spacex"), path))
    finally:
        eng.set_key(k, orig)


def test_v018_committed_unchanged_by_q3(eng):
    """R5f：Q3 無新股權——2026 已到位融資對 S-1/A 前九月特別股收入 110.432 差 ≤ 0.5；R5e 9-30 現金差額拆解四項相加＝差。"""
    s, i = _ids()
    com = eng.get_name("FND_Committed")[1]
    assert abs(com - eng.get_name(s["fund_s1a_pref_2026_9m"])) <= eng.get_name(i["s1f_cmt_tol"])
    assert abs(eng.get_name("FND_CommittedS1aGap") - (com - eng.get_name(s["fund_s1a_pref_2026_9m"]))) < 1e-9


def test_v018_charity_not_in_per_person_base(eng):
    """v0.1.8 核對後：S-1/A 下半年非現金捐贈（1.960 × 2）不進 2026 每人股權報酬基準（不沿用到 2027–2030）；2026 全成本另計（判斷題 s1a_chr_q4_switch：1＝Q4 也計，0＝只計 Q3）。"""
    s, i = _ids()
    g = lambda k: eng.get_name(s[k])  # noqa: E731
    chr_ = g("charity_s1a_2026q3")
    sbc, chc, proxy = eng.get_name("COST_SBC"), eng.get_name("COST_S1aCharityInCost"), eng.get_name("COST_S1fSBCProxy2026")
    assert abs(chc[1] - 2 * chr_) < 1e-9 and all(v == 0 for k, v in enumerate(chc) if k != 1)
    assert abs((sbc[1] - chc[1]) - (proxy - 2 * chr_)) < 1e-9          # 每人基準 × 人數＝代理 − 捐贈
    assert abs(sbc[1] - proxy) < 1e-9                                   # 基準（Q4 亦計）2026 當年仍含捐贈
    from engine import resolve_key
    k = i["s1a_chr_q4_switch"]
    orig = eng.get(*resolve_key(eng.names, k))
    later = sbc[2:]
    try:
        eng.set_key(k, 0)
        s2 = eng.get_name("COST_SBC")
        assert abs(s2[1] - (proxy - chr_)) < 1e-9 and all(abs(a - b) < 1e-12 for a, b in zip(s2[2:], later))
    finally:
        eng.set_key(k, orig)


def test_v019_tk_back_reproduces_v018(eng):
    """v0.1.9：TK 快照回到 v21.7（TK_Link 全部具名值取自 v0.1.8 Excel）時，v0.1.8 的全部計算具名範圍逐位相同（公式未改；差異只來自 TK）。"""
    n_cmp = _reproduce(eng, V018_XLSX, _tk_ov(eng, V018_XLSX))
    assert n_cmp > 1700


def test_v019_contract_pay_and_cash_unchanged(eng):
    """v0.1.9 工作單：合約付款（GW × 持有比）應不變；TK 產出比只改 VR 等值換算與 η，現金與命題 2 輸出逐位等於 v0.1.8。"""
    wo = openpyxl.load_workbook(V018_XLSX)
    wv = openpyxl.load_workbook(V018_XLSX, data_only=True)
    names = [n for n in wo.defined_names if n.startswith("COST_Pay_")] + [
        "COST_ContractPay", "COST_Compute", "CMP_SupplyGW", "REV_NetCapped", "FND_FCF", "FND_CashNoExt", "FND_CashNoExtTrough", "FND_ExtNeedCum",
        "FND_MinCash", "FND_FirstGapYear", "COST_Coverage", "COST_GapPosSustainYear"]
    for n in names:
        sh, ref = wo.defined_names[n].attr_text.split("!")
        cells = wv[sh][ref.replace("$", "")]
        want = [c.value for row in cells for c in row] if isinstance(cells, tuple) else [cells.value]
        got = eng.get_name(n)
        got = got if isinstance(got, list) else [got]
        for a, b in zip(got, want):
            if isinstance(b, (int, float)) and not isinstance(b, bool):
                assert abs(a - b) <= 1e-9 * max(1, abs(b)), (n, a, b)
            else:
                assert (a in (None, "") and b in (None, "")) or a == b, (n, a, b)

"""OpenAI v0.6.5 PL 頁（標準損益表；個股 HTML 版面規格 v1.3 D 節＋J 節「未上市、無市場價格的模型」）。

只列示，不被任何計算頁引用（Checks 除外）；既有工作表數值不變。期間 FY2023–FY2030（歷史 3 年＋預測 5 年），單位十億美元。
- 歷史：只放有公司（外流經查核財報）揭露者——FY2024 營收（SRC_OAI_001）、FY2025 營收、營業成本、研發、銷售、管理費用（SRC_OAI_002／003／087／129／130）。
  未揭露的格寫「未揭露」，不補編（規格 J 節）。FY2025 稅後淨利等 SRC 沒有，同樣「未揭露」。
- 預測（FY2026–FY2030，本模型基準情境）：營收＝營收總額（截頂後）；營業成本＝推論算力成本；營業費用＝研發算力＋閒置算力＋非算力成本（含股權報酬）＋Microsoft 營收分成。
  營業利益因此等於 Cost 頁「差額（含股權報酬；現金口徑）」（核對列，應為 0）。其他收入、稅前、所得稅、稅後淨利、EPS 本模型未建（不補假設）：寫「未建模」。
- 口徑：GAAP（外流財報為 GAAP 認列）；預測的算力成本取合約付款（現金口徑），與 GAAP 租賃認列有時間差；Microsoft 分成在 GAAP 的列示位置未揭露，本表放營業費用。
"""
from __future__ import annotations

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_SEC, header, put, title

YEARS = list(range(2023, 2031))
COLS = list("CDEFGHIJ")            # FY2023–FY2030
NOT_DISC, NOT_MODEL, NA = "未揭露", "未建模", "—"


def build_p8(wb, S, nm):
    ws = wb.create_sheet("PL")
    title(ws, "PL — 標準損益表（GAAP；FY2023–FY2030；十億美元）",
          "歷史取外流經查核財報（SRC_OAI，二手轉載）；未揭露者留空標「未揭露」。預測取本模型基準情境；其他收入、稅前、所得稅、稅後淨利、EPS 本模型未建，標「未建模」。",
          "只列示，不被任何計算頁引用（Checks 除外）。藍字＝輸入，黑字＝公式。營業利益預測列＝Cost 頁差額（含股權報酬；現金口徑），核對列見下。")
    header(ws, 5, ["項目", "單位"] + [f"FY{y}{'（預測）' if y >= 2026 else ''}" for y in YEARS] + ["說明"])
    R = {}
    r = [6]
    fmt = "#,##0.00"

    def row(key, label, vals, unit="$B", note="", f=fmt, bold=False):
        n = r[0]
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for c, v in zip(COLS, vals):
            if v is None:
                continue
            is_f = isinstance(v, str) and v.startswith("=")
            put(ws, f"{c}{n}", v, F_CALC, fmt=None if (isinstance(v, str) and not is_f) else f)
        if note:
            put(ws, f"K{n}", note, F_NOTE)
        R[key] = n
        r[0] += 1
        return n

    m = lambda k: k + 1                       # 預測年 2026–2030 對應模型欄索引（D–I＝2025–2030）：2026＝2  # noqa: E731
    fc = list(range(2, 7))
    cost = lambda rw, k: f"Cost!${'DEFGHI'[k - 1]}${rw}"  # noqa: E731  Cost 頁固定列：73 研發算力、74 推論算力、75 閒置算力
    n = r[0]
    row("rev", "營收", [NOT_DISC, f"={S('營收：FY2024')}", f"={S('營收：FY2025')}"] + [f"=INDEX(REV_GrossCapped,1,{k})" for k in fc], note="FY2024／FY2025：SRC_OAI_001／002（外流財報轉載）；預測：營收總額（截頂後）", bold=True)
    n = R["rev"]
    row("yoy", "營收年增率（YoY）", [NA, NA] + [f"={c}{n}/{p}{n}-1" for p, c in zip("DEFGHI", "EFGHIJ")], "比例", f="0.0%")
    row("cogs", "營業成本（COGS）", [NOT_DISC, NOT_DISC, f"={S('營業成本（cost of revenue）：FY2025')}"] + [f"={cost(74, k)}" for k in fc],
        note="FY2025：外流財報（含推論）；預測：推論算力成本（Cost 頁推論算力；2025 模型推論 8.4 對財報 7.5，口徑差見 Checks）")
    c_ = R["cogs"]
    row("gp", "毛利", [NOT_DISC, NOT_DISC] + [f"={c}{n}-{c}{c_}" for c in "EFGHIJ"], bold=True)
    g_ = R["gp"]
    row("gpm", "毛利率", [NA, NA] + [f"={c}{g_}/{c}{n}" for c in "EFGHIJ"], "比例", f="0.0%")
    row("opex", "營業費用", [NOT_DISC, NOT_DISC, f"={S('研發費用總額：2025')}+{S('銷售費用（sales and marketing）：FY2025')}+{S('管理費用（general and administrative）：FY2025')}"]
        + [f"={cost(73, k)}+{cost(75, k)}+INDEX(COST_NonCompInclSBC,1,{k})+INDEX(REV_MSShareCapped,1,{k})" for k in fc],
        note="FY2025：研發＋銷售＋管理（外流財報）；預測：研發算力＋閒置算力＋非算力成本（含股權報酬）＋Microsoft 營收分成")
    o_ = R["opex"]
    row("oi", "營業利益", [NOT_DISC, NOT_DISC] + [f"={c}{g_}-{c}{o_}" for c in "EFGHIJ"], bold=True)
    i_ = R["oi"]
    row("oim", "營業利益率", [NA, NA] + [f"={c}{i_}/{c}{n}" for c in "EFGHIJ"], "比例", f="0.0%")
    row("oth", "其他收入（淨額）", [NOT_DISC] * 3 + [NOT_MODEL] * 5, note="外流財報摘要未含；本模型未建利息與公允價值變動")
    row("pbt", "稅前利益", [NOT_DISC] * 3 + [NOT_MODEL] * 5, bold=True)
    row("tax", "所得稅", [NOT_DISC] * 3 + [NOT_MODEL] * 5)
    row("ni", "稅後淨利", [NOT_DISC] * 3 + [NOT_MODEL] * 5, bold=True, note="預測的自由現金流見 Funding 頁；本模型不做稅後淨利")
    row("nim", "淨利率", [NA] * 8, "比例")
    row("eps", "EPS（完全稀釋，美元）", [NOT_DISC] * 3 + ["不適用"] * 5, "$", note="未上市、無評價股數；不適用")
    row("niyoy", "淨利年增率", [NA] * 8, "比例")
    r[0] += 1
    put(ws, f"A{r[0]}", "核對（只列示）", F_BOLD, fill=FILL_SEC)
    r[0] += 1
    row("chk_fc", "核對：營業利益（預測）− Cost 頁差額（含股權報酬；現金口徑）（應為 0）", [None, None, None] + [f"={c}{i_}-INDEX(COST_GapCash,1,{k})" for c, k in zip("FGHIJ", fc)])
    row("chk_25", "核對：FY2025 營業利益 − SRC_OAI_131 營業損失（外流財報；差為四捨五入）", [None, None, f"=E{i_}+{S('營業損失（loss from operations）：FY2025')}"], note="SRC 營業損失 20.92（正值＝虧損）對 13.07−7.5−26.48＝−20.91：差 0.01 為四捨五入")
    for nm_, key in (("PL_Rev", "rev"), ("PL_RevYoY", "yoy"), ("PL_COGS", "cogs"), ("PL_GP", "gp"), ("PL_GPM", "gpm"), ("PL_Opex", "opex"), ("PL_OI", "oi"),
                     ("PL_OIM", "oim"), ("PL_Oth", "oth"), ("PL_PBT", "pbt"), ("PL_Tax", "tax"), ("PL_NI", "ni"), ("PL_NIM", "nim"), ("PL_EPS", "eps"), ("PL_NIYoY", "niyoy")):
        nm(wb, nm_, f"PL!$C${R[key]}:$J${R[key]}")
    nm(wb, "PL_ChkFc", f"PL!$F${R['chk_fc']}:$J${R['chk_fc']}")
    nm(wb, "PL_Chk25", f"PL!$E${R['chk_25']}")
    ws.column_dimensions["A"].width = 64
    ws.column_dimensions["B"].width = 7
    for c in COLS:
        ws.column_dimensions[c].width = 13
    ws.column_dimensions["K"].width = 90
    ws.freeze_panes = "C6"
    return dict(R=R)


def p8_checks(P8):
    return [
        ("v0.6.5 PL：營業利益（預測）對 Cost 差額的絕對差之和（應為 0）", "=SUMPRODUCT(ABS(PL_ChkFc))", 0, "tol", "PL 頁只列示；營業利益預測＝營收總額−推論算力−（研發算力＋閒置算力＋非算力含股權報酬＋Microsoft 分成）"),
        ("v0.6.5 PL：FY2025 營業利益對外流財報營業損失的差（絕對值 ≤ 0.05）", "=ABS(PL_Chk25)", 0.05, "le", "四捨五入差 0.01"),
    ]

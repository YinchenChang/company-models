# 英特爾 v0.1.0 展示層：README 重點觀察 2、3（公式帶數字；重點觀察 1 由 tools/relative_multiple 相對倍數段寫入）
def _lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def _F(*parts):
    return "=" + "&".join(_lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def apply(wb, G):
    RD = wb["README"]; PB = G["PB"]["B"]; SR = G["SENS_ROWS"]
    T = lambda s: ("t", s)
    tx = lambda ref, f: f'TEXT({ref},"{f}")'
    SP = "+0.0%;-0.0%"
    r0 = G["OBS_ROW0"]
    RD.cell(r0 + 1, 1).value = _F(
        T("重點觀察 2：Foundry 虧損收斂與資本支出決定獲利。FY2026 Foundry 營業虧損 "), tx(f"-PL!D{PB['fndoi']}", "0.0"), T(" 十億美元（營業利益率 "), tx(f"PL!D{PB['fm']}", "0%"),
        T("），情境 B 到 FY2029 收斂到 "), tx(f"PL!G{PB['fm']}", "0%"), T("（A "), tx(f"PL!G{G['PB']['A']['fm']}", "0%"), T("、C "), tx(f"PL!G{G['PB']['C']['fm']}", "0%"),
        T("）；Foundry 利益率 ±5 個百分點使加權目標價 "), tx(f"Sens!J{SR[('fm_adj', 'lo')]}", SP), T("／"), tx(f"Sens!J{SR[('fm_adj', 'hi')]}", SP),
        T("。總資本支出 FY2026 "), tx(f"PL!D{PB['cx']}", "0.0"), T("、FY2029 "), tx(f"PL!G{PB['cx']}", "0.0"), T(" 十億美元（占營收 "), tx(f"PL!D{PB['cxs']}", "0%"), T(" → "), tx(f"PL!G{PB['cxs']}", "0%"),
        T("），公司口徑未扣政府補助與夥伴挹注。"))
    RD.cell(r0 + 2, 1).value = _F(
        T("重點觀察 3：AI 與 CPU 成長有限且已被反映。DCAI 營收 FY2026 "), tx(f"PL!D{PB['dcr']}", "0.0"), T(" → FY2029 "), tx(f"PL!G{PB['dcr']}", "0.0"), T(" 十億美元（情境 B；AI 相關 "), tx(f"PL!G{PB['dai']}", "0.0"),
        T("）；DCAI 中 AI 相關營收 ±30% 使目標價 "), tx(f"Sens!J{SR[('ai_adj', 'lo')]}", SP), T("／"), tx(f"Sens!J{SR[('ai_adj', 'hi')]}", SP),
        T("，DCAI 營業利益率 ±3 個百分點 "), tx(f"Sens!J{SR[('dcm_adj', 'lo')]}", SP), T("／"), tx(f"Sens!J{SR[('dcm_adj', 'hi')]}", SP),
        T("；CCPG（客戶端）FY2029 營收 "), tx(f"PL!G{PB['ccg']}", "0.0"), T(" 十億美元，年增僅 1–2%。"))

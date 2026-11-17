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
    apply2(wb, G)


def apply2(wb, G):
    """README 展示層（命題摘要、命題分項、白話重點）與命題回覆（relayout 讀取）；數字皆公式。"""
    RD = wb["README"]; V = wb["Val"]; PB = G["PB"]["B"]; PA = G["PB"]["A"]; PC = G["PB"]["C"]
    T = lambda s: ("t", s)
    tx = lambda ref, f: f'TEXT({ref},"{f}")'
    SP = "+0.0%;-0.0%"

    def vl(prefix, col="B"):
        for r in range(1, V.max_row + 1):
            if str(V.cell(r, 1).value or "").strip().startswith(prefix) and V.cell(r, 2).value is not None:
                return f"Val!{col}{r}"
        raise KeyError(prefix)
    tp, dcf, cadj, up, rec = vl("目標價（DCF 與 Comparable 加權）"), vl("DCF 目標價（情境加權）"), vl("Comparable 目標價（相對倍數調整後）"), vl("上檔空間：加權目標價"), vl("投資建議")
    r = RD.max_row + 2
    RD.cell(r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.1.0 起；上方各列原文保留給下游）"); r += 1
    RD.cell(r, 1, "命題：英特爾的獲利由 Intel Foundry 的虧損收斂速度與資本支出強度決定，Intel Products（客戶端與資料中心 CPU）提供現金流；以下三個問題依序回答。"); r += 1
    RD.cell(r, 1, "命題 1：Foundry 虧損能收斂多快。外部客戶（18A／14A、先進封裝）能把外部代工營收帶到多大，Foundry 營業利益率能否在 FY2029 前接近損益兩平。"); r += 1
    RD.cell(r, 1, "命題 2：Intel Products 能撐多少。客戶端（CCPG）接近成熟、資料中心（DCAI）受 AI 與伺服器 CPU 需求帶動，兩者的利潤能否支付 Foundry 虧損與資本支出。"); r += 1
    RD.cell(r, 1, "命題 3：每股值多少。用 DCF 與同業倍數（含英特爾相對同業的長期倍數比）算出的目標價，和現價相比是高是低。"); r += 1
    for k in range(3):
        RD.cell(r, 1, f"=A{G['OBS_ROW0'] + k}"); r += 1
    r += 1
    RD.cell(r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上面三個命題）"); r += 1
    RD.cell(r, 1).value = _F(
        T("Foundry 虧損能收斂多快？答：緩慢，而且是假設。FY2026 Foundry 營業利益率 "), tx(f"PL!D{PB['fm']}", "0%"), T("，FY2029 情境 A／B／C 為 "), tx(f"PL!G{PA['fm']}", "0%"), T("／"), tx(f"PL!G{PB['fm']}", "0%"), T("／"), tx(f"PL!G{PC['fm']}", "0%"),
        T("；外部代工營收 FY2029 為 "), tx(f"PL!G{PA['ext']}", "0.0"), T("／"), tx(f"PL!G{PB['ext']}", "0.0"), T("／"), tx(f"PL!G{PC['ext']}", "0.0"), T(" 十億美元。基準情境到 FY2029 仍虧損 "), tx(f"-PL!G{PB['fndoi']}", "0.0"), T(" 十億美元。")); r += 1
    RD.cell(r, 1).value = _F(
        T("Intel Products 能撐多少？答：獲利夠付利息但不夠付 Foundry。FY2026 CCPG 與 DCAI 營業利益合計 "), tx(f"PL!D{PB['ccoi']}+PL!D{PB['dcoi']}", "0.0"), T(" 十億美元，FY2029 情境 B 為 "), tx(f"PL!G{PB['ccoi']}+PL!G{PB['dcoi']}", "0.0"),
        T("；同期 Foundry 虧損 "), tx(f"-PL!G{PB['fndoi']}", "0.0"), T("、Corporate 未分攤費用 "), tx(f"-PL!G{PB['cu']}", "0.0"), T(" 十億美元，資本支出 "), tx(f"PL!G{PB['cx']}", "0.0"), T(" 十億美元。")); r += 1
    RD.cell(r, 1).value = _F(
        T("每股值多少？答：加權目標價 "), tx(tp, "#,##0.0"), T(" 美元（DCF "), tx(dcf, "#,##0.0"), T("、Comparable 相對倍數調整後 "), tx(cadj, "#,##0.0"), T("，權重各半），比現價 "), tx("px", "#,##0.0"), T(" 美元"),
        f'IF({up}<0,"低 ","高 ")&TEXT(ABS({up}),"0.0%")', T("，投資建議"), rec, T("。")); r += 2
    RD.cell(r, 1, "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間（HTML 估值表用；加權目標價使用的 Comparable 即此值；投資建議橫幅中間格放只看 Comparable 的未調整值）")
    RD.cell(r, 2, f"={cadj}/px-1").number_format = "0.0%"

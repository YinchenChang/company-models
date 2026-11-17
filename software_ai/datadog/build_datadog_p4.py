# build_datadog.py 的第四部分（Recon、Sens、Sources、README、列號對照、存檔）；由 build_datadog_p3.py exec。
DIAG_TEXT = open(os.path.join(HERE, "data", "diagnosis.txt"), encoding="utf-8").read().strip() if os.path.exists(os.path.join(HERE, "data", "diagnosis.txt")) else "診斷尚未撰寫"
SENSJ = SENS or {}

# ====================================================================================================== Recon
RC["A1"] = "Recon — 公司指引、市場共識、校準與合理性關卡（公司指引與共識＝Interested-party，只作對照，不進基準）"
RC["A2"] = ("指引來源：2026Q3 財報新聞稿（2026-11-05 前後）。公司指引為保守口徑：近三季營收皆超出指引上緣 3.8%–4.7%（第 5 節）。"
            "市場共識：FY2027 營收成長 23%（財報後，TIKR）、FY2027 EPS 與營收的財報前值；平均目標價 276.92（TipRanks，2026-11-12）。口徑：公司 non-GAAP 排除股票薪酬，本模型 DCF 以 GAAP 為底。")
RCR = {}
r = 4
RC.cell(r, 1, "1. 公司指引對模型情境 B").font = BOLD; r += 1
hdr(RC, r, ["項目", "期間", "模型情境 B", "指引下緣", "指引上緣", "模型 ÷ 指引中點 − 1", "說明"]); r += 1
PB_ = P_["B"]
G1 = [("q4rev", "營收", "Q4 26E", f"=PL!G{PB_['rev']}", "guid.g_q4lo", "guid.g_q4hi", "#,##0.0", "模型 Q4 季增 q4_g 高於指引上緣，依近三季超出指引 3.8%–4.7% 的紀錄；指引已假設最大客戶以最低承諾計"),
      ("fyrev", "營收", "FY2026E", f"=PL!H{PB_['rev']}", "guid.g_fylo", "guid.g_fyhi", "#,##0.0", "9M26 實績＋Q4 26E"),
      ("q4oi", "營業利益（non-GAAP）", "Q4 26E", f"=PL!G{PB_['oin']}", "guid.g_ng_q4lo", "guid.g_ng_q4hi", "#,##0.0", "模型 non-GAAP 費用率持平於近四季；指引中點 285"),
      ("fyoi", "營業利益（non-GAAP）", "FY2026E", f"=PL!H{PB_['oin']}", "guid.g_ng_fylo", "guid.g_ng_fyhi", "#,##0.0", ""),
      ("fyeps", "EPS（non-GAAP，稀釋，US$）", "FY2026E", f"=PL!H{PB_['ngeps']}", "guid.g_ngeps_fylo", "guid.g_ngeps_fyhi", "#,##0.00", "公司稀釋股數假設約 375；模型 9M26 實績加權與 Q4 假設見 PL 頁")]
for k, lab, per, f, lo, hi, fm, note in G1:
    RCR[k] = r
    RC.cell(r, 1, lab); RC.cell(r, 2, per); RC.cell(r, 3, f).number_format = fm
    for cc, fid in ((4, lo), (5, hi)):
        c = RC.cell(r, cc, fv(fid)); c.font = BLUE; c.number_format = fm
    RC.cell(r, 6, f"=C{r}/AVERAGE(D{r}:E{r})-1").number_format = "+0.0%;-0.0%"; RC.cell(r, 7, note + "（Interested-party；" + F[lo]["source"][:60] + "）")
    r += 1
r += 1
RC.cell(r, 1, "2. 市場共識對模型（只作對照）").font = BOLD; r += 1
hdr(RC, r, ["項目", "期間", "模型情境 A", "模型情境 B", "模型情境 C", "共識", "說明"]); r += 1
RCR["g27"] = r
RC.cell(r, 1, "營收年增率"); RC.cell(r, 2, "FY2027E")
for cc, s in ((3, "A"), (4, "B"), (5, "C")): RC.cell(r, cc, f"=PL!I{P_[s]['yoy']}").number_format = "0.0%"
c = RC.cell(r, 6, fv("cons.rev27g")); c.font = BLUE; c.number_format = "0.0%"; RC.cell(r, 7, "TIKR（財報後；Interested-party：賣方預估）"); r += 1
RCR["rev27"] = r
RC.cell(r, 1, "營收（US$M）；共識為財報前值（已過時）"); RC.cell(r, 2, "FY2027E")
for cc, s in ((3, "A"), (4, "B"), (5, "C")): RC.cell(r, cc, f"=PL!I{P_[s]['rev']}").number_format = "#,##0.0"
c = RC.cell(r, 6, fv("cons.rev27.pre")); c.font = BLUE; c.number_format = "#,##0.0"; RC.cell(r, 7, "2026-10-26 財報前共識；財報後更新值未找到"); r += 1
RCR["eps27"] = r
RC.cell(r, 1, "EPS（non-GAAP，稀釋，US$）；共識為財報前值"); RC.cell(r, 2, "FY2027E")
for cc, s in ((3, "A"), (4, "B"), (5, "C")): RC.cell(r, cc, f"=PL!I{P_[s]['ngeps']}").number_format = "#,##0.00"
c = RC.cell(r, 6, fv("cons.eps27.pre")); c.font = BLUE; c.number_format = "0.00"; RC.cell(r, 7, "2026-10-26 財報前共識；財報後更新值未找到"); r += 1
RCR["tpc"] = r
RC.cell(r, 1, "加權目標價 vs 賣方平均目標價（高 335、低 171；家數頁面未載）"); RC.cell(r, 2, "目標時點"); RC.cell(r, 3, f"=Val!B{VR['tp']}").number_format = "#,##0.00"
RC.cell(r, 4, "=cons_tp").number_format = "#,##0.00"; RC.cell(r, 5, f"=C{r}/D{r}-1").number_format = "+0.0%;-0.0%"; RC.cell(r, 7, "共識目標價日期 2026-11-12；賣方誘因：評等與投行業務；只作對照"); r += 2
RC.cell(r, 1, "3. AI 原生客戶占營收比交叉檢查（Derived；公司自 2025Q3 後不再揭露）").font = BOLD; r += 1
hdr(RC, r, ["項目", "期間", "數值", "", "", "", "說明"]); r += 1
RCR["ai3q25"] = r; RC.cell(r, 1, "AI 原生客戶占營收比（公司揭露）"); RC.cell(r, 2, "2025Q3"); c = RC.cell(r, 3, fv("ai.sh.3q25")); c.font = BLUE; c.number_format = "0.0%"; RC.cell(r, 7, fs("ai.sh.3q25") + "（Interested-party）"); r += 1
RCR["nonai"] = r; RC.cell(r, 1, "非 AI 客戶營收年增（法說 high 20s 取 28%）"); RC.cell(r, 2, "2026Q3"); c = RC.cell(r, 3, fv("ai.nonai.3q26")); c.font = BLUE; c.number_format = "0.0%"; RC.cell(r, 7, fs("ai.nonai.3q26") + "（Interested-party；high 20s 取點為 Assumed）"); r += 1
RCR["ai3q26"] = r; RC.cell(r, 1, "反推 2026Q3 AI 原生客戶占營收比＝1 − (1 − 2025Q3 占比) × (1 ＋ 非 AI 年增) × 3Q25 營收 ÷ 3Q26 營收"); RC.cell(r, 2, "2026Q3")
RC.cell(r, 3, f"=1-(1-C{RCR['ai3q25']})*(1+C{RCR['nonai']})*q_rev_3q25/q_rev_3q26").number_format = "0.0%"; RC.cell(r, 7, "Derived；與 In 頁 sa_26（FY2026E 平均占比）對照"); r += 1
RCR["sa26"] = r; RC.cell(r, 1, "In 頁 sa_26（FY2026E 全年平均占比）"); RC.cell(r, 3, "=sa_26").number_format = "0.0%"; r += 2
RC.cell(r, 1, "4. 公司指引的保守程度（歷史：實績超出指引上緣）").font = BOLD; r += 1
hdr(RC, r, ["項目", "期間", "超出幅度", "", "", "", "說明"]); r += 1
for q, fid in (("2026Q1", "guid.beat1"), ("2026Q2", "guid.beat2"), ("2026Q3", "guid.beat3")):
    RC.cell(r, 1, "營收實績超出該季指引上緣"); RC.cell(r, 2, q); c = RC.cell(r, 3, fv(fid)); c.font = BLUE; c.number_format = "0.0%"; RC.cell(r, 7, F[fid]["source"][:80] + "（Derived）"); r += 1
RCR["beat_end"] = r - 1
r += 1
RC.cell(r, 1, "5. 每任務 token 校準（G2 第 1 條）：公司揭露推得值 vs Tokenomics 聊天基數 × 任務倍數").font = BOLD; r += 1
hdr(RC, r, ["項目", "單位", "數值", "", "", "", "說明"]); r += 1
RCR["tk_d"] = r; RC.cell(r, 1, "每任務 token（由公司揭露校準：成本占比 × 每任務 credit 收入 ÷ 校準價格）"); RC.cell(r, 2, "百萬 token"); RC.cell(r, 3, "=tk_tokens").number_format = "0.000"; RC.cell(r, 7, "In 頁 tk_tokens（Derived／Interested-party）：基準值"); r += 1
RCR["tk_c"] = r; RC.cell(r, 1, "對照：聊天基數 × 8 小時任務倍數（IF_TaskTokMult 末項）"); RC.cell(r, 2, "百萬 token")
c = RC.cell(r, 3, 0.003 * tkv("IF_TaskTokMult")[-1]); c.font = BLUE; c.number_format = "0.000"; RC.cell(r, 7, TKS + "；聊天基數 3,000 token × 倍數 498（Analogy；不作基準）"); r += 1
RCR["tk_r"] = r; RC.cell(r, 1, "揭露校準值 ÷ 聊天基數類推值"); RC.cell(r, 3, f"=C{RCR['tk_d']}/C{RCR['tk_c']}").number_format = "0.00"; RC.cell(r, 7, "差距大小只作對照；成本方向的不確定性由 Sens（token 價格、路由、聯動）涵蓋"); r += 1
RCR["tk_cost"] = r; RC.cell(r, 1, "每任務推論成本（US$）：FY2027E 情境 B（Thesis 頁）"); RC.cell(r, 3, f"=Thesis!I{T['B']['ctask']}").number_format = "0.000"; RC.cell(r, 7, "token 價格只讀 Tokenomics 快照"); r += 1
RCR["tk_rev"] = r; RC.cell(r, 1, "每任務 credit 收入（US$）：FY2027E 情境 B"); RC.cell(r, 3, f"=cr_task*Thesis!I{T['B']['crp']}").number_format = "0.000"; RC.cell(r, 7, "耗用 credits × 每 credit 價格"); r += 2
RC.cell(r, 1, "6. 反解（G2 第 9 條）：主導變數整本重算，使加權目標價等於現價所需的值").font = BOLD; r += 1
hdr(RC, r, ["項目", "單位", "數值", "", "", "", "說明"]); r += 1
RVS = SENSJ.get("reverse") or {}
RCR["sv_var"] = r; RC.cell(r, 1, "主導變數：股票薪酬占營收每年變動 sbc_st（歷史規模經濟的延伸斜率；DCF 把股票薪酬視為費用，故承重最大）"); RC.cell(r, 2, "%"); RC.cell(r, 3, "=sbc_st").number_format = "0.00%"; RC.cell(r, 7, "基準＝歷史斜率"); r += 1
RCR["sv_x"] = r; RC.cell(r, 1, "支持現價所需值（整本重算，割線法；Derived）"); RC.cell(r, 2, "%"); c = RC.cell(r, 3, RVS.get("x")); c.number_format = "0.00%"; RC.cell(r, 7, RVS.get("note", "sens_datadog.py 逐次整本以 LibreOffice 重算")); r += 1
RCR["sv_tp"] = r; RC.cell(r, 1, "代入所需值後的加權目標價（整本重算）與現價"); RC.cell(r, 2, "US$"); c = RC.cell(r, 3, RVS.get("tp")); c.number_format = "#,##0.00"; RC.cell(r, 4, "=px").number_format = "#,##0.00"; RC.cell(r, 5, f'=IF(C{r}="","",C{r}/D{r}-1)').number_format = "+0.0%;-0.0%"; RC.cell(r, 7, "差距應接近零；若所需值超出合理範圍（例如 SBC 占營收低於下限），代表現價無法由該變數單獨支持"); r += 1
RCR["sv_sbcfloor"] = r; RC.cell(r, 1, "參考：該所需值對應 FY2032E 股票薪酬占營收（＝MAX(下限, 近四季 ＋ 6 × 所需值)）"); RC.cell(r, 2, "%"); RC.cell(r, 3, f'=IF(C{RCR["sv_x"]}="","",MAX(sbc_floor,sbc_0+6*C{RCR["sv_x"]}))').number_format = "0.0%"; r += 2
RC.cell(r, 1, "7. 合理性關卡（G3 第 3 條；發布前必查）").font = BOLD; r += 1
hdr(RC, r, ["項目", "判定值", "門檻", "是否觸發", "", "", "說明"]); r += 1
RCR["gt1"] = r; RC.cell(r, 1, "DCF 目標價 ÷ 現價"); RC.cell(r, 2, f"=Val!B{VR['tpdcf']}/px").number_format = "0.00"; RC.cell(r, 3, "< 0.25 或 > 4"); RC.cell(r, 4, f'=IF(OR(B{r}<0.25,B{r}>4),"觸發","未觸發")'); RC.cell(r, 7, "觸發時須診斷"); r += 1
RCR["gt2"] = r; RC.cell(r, 1, "EV 為負（三情境最小 EV，US$M）"); RC.cell(r, 2, f"=MIN(Val!B{VR['ev']}:D{VR['ev']})").number_format = "#,##0"; RC.cell(r, 3, "< 0"); RC.cell(r, 4, f'=IF(B{r}<0,"觸發","未觸發")'); r += 1
RCR["gt3"] = r; RC.cell(r, 1, "評等（估值評等）"); RC.cell(r, 2, f"=Val!B{VR['valrate']}"); RC.cell(r, 3, "強力放空或強力買進"); RC.cell(r, 4, f'=IF(OR(B{r}="強力放空",B{r}="強力買進"),"觸發","未觸發")'); r += 1
RCR["gt4"] = r; RC.cell(r, 1, "模型 FY2026E 營收 ÷ 指引中點 − 1"); RC.cell(r, 2, f"=F{RCR['fyrev']}").number_format = "+0.0%;-0.0%"; RC.cell(r, 3, "絕對值 > 10%"); RC.cell(r, 4, f'=IF(ABS(B{r})>0.1,"觸發","未觸發")'); RC.cell(r, 7, "公司只給 FY2026 與 Q4 指引；FY2027 無公司指引（未找到），改以共識成長 23% 對照第 2 節"); r += 1
RCR["gt5"] = r; RC.cell(r, 1, "模型 FY2026E non-GAAP 營業利益 ÷ 指引中點 − 1"); RC.cell(r, 2, f"=F{RCR['fyoi']}").number_format = "+0.0%;-0.0%"; RC.cell(r, 3, "絕對值 > 10%"); RC.cell(r, 4, f'=IF(ABS(B{r})>0.1,"觸發","未觸發")'); r += 1
RCR["gt6"] = r; RC.cell(r, 1, "模型 FY2027E 營收成長（情境 B）與共識成長之差（百分點）"); RC.cell(r, 2, f"=(D{RCR['g27']}-F{RCR['g27']})*100").number_format = "+0.0;-0.0"; RC.cell(r, 3, "絕對值 > 10 個百分點"); RC.cell(r, 4, f'=IF(ABS(B{r})>10,"觸發","未觸發")'); r += 1
RCR["gt_n"] = r; RC.cell(r, 1, "觸發項目數"); RC.cell(r, 2, f'=COUNTIF(D{RCR["gt1"]}:D{RCR["gt6"]},"觸發")').number_format = "0"; r += 1
RCR["diag"] = r; RC.cell(r, 1, "診斷結論（寫入 README；來源 data/diagnosis.txt）"); RC.cell(r, 2, DIAG_TEXT); r += 2
RC.cell(r, 1, "7b. DCF 與公司自由現金流對帳（診斷用：確認 DCF 偏低是口徑與折現率，不是現金流算錯）").font = BOLD; r += 1
hdr(RC, r, ["項目", "期間", "US$M", "", "", "", "說明"]); r += 1
CB_ = C_["B"]; DB_ = DR_["B"]
RCR["rc_fcf"] = r; RC.cell(r, 1, "自由現金流（公司定義，CF 頁情境 B）"); RC.cell(r, 2, "FY2027E"); RC.cell(r, 3, f"=CF!I{CB_['fcf']}").number_format = "#,##0.0"; r += 1
RCR["rc_int"] = r; RC.cell(r, 1, "− 利息收入淨額 ×(1 − 稅率)（FCFF 不含利息，現金另加在股權價值）"); RC.cell(r, 2, "FY2027E"); RC.cell(r, 3, f"=PL!I{PB_['onet']}*(1-PL!I{PB_['taxr']})").number_format = "#,##0.0"; r += 1
RCR["rc_fcffs"] = r; RC.cell(r, 1, "＝FCFF（加回全額股票薪酬；DCF 頁）"); RC.cell(r, 2, "FY2027E"); RC.cell(r, 3, f"=DCF!I{DB_['fcffs']}").number_format = "#,##0.0"; r += 1
RCR["rc_chk"] = r; RC.cell(r, 1, "檢查：公司定義自由現金流 − 利息稅後 − FCFF（應為 0）"); RC.cell(r, 3, f"=C{RCR['rc_fcf']}-C{RCR['rc_int']}-C{RCR['rc_fcffs']}").number_format = "0.000"; r += 1
RCR["rc_sbc"] = r; RC.cell(r, 1, "FCFF（股票薪酬視為費用；主值）與加回股票薪酬之差＝股票薪酬"); RC.cell(r, 2, "FY2027E"); RC.cell(r, 3, f"=DCF!I{DB_['fcffs']}-DCF!I{DB_['fcff']}").number_format = "#,##0.0"; RC.cell(r, 7, "DCF 偏低的第一個原因：股票薪酬占公司定義自由現金流的大部分"); r += 1
RCR["rc_wacc"] = r; RC.cell(r, 1, "WACC／終值折現年數／終值占 EV（情境 B）"); RC.cell(r, 3, f"=Val!B{VR['wacc']}").number_format = "0.00%"; RC.cell(r, 4, "=t_tv").number_format = "0.00"; RC.cell(r, 5, f"=Val!C{VR['tvsh']}").number_format = "0%"; RC.cell(r, 7, "第二個原因：折現率約 11.8%；第三個：FY2032 營收年增仍約 8% 即降到永續成長率"); r += 1
RCR["rc_psv_s"] = r; RC.cell(r, 1, "DCF（股票薪酬加回）每股價值（US$，情境加權，評價日；只作對照）"); RC.cell(r, 3, f"=Val!B{VR['psv_s']}").number_format = "#,##0.00"; RC.cell(r, 7, "對照用，不進加權目標價（G5 第 3 條）"); r += 1
RCR["rc_tp_s"] = r; RC.cell(r, 1, "DCF（股票薪酬加回）目標價（US$，情境加權；只作對照）"); RC.cell(r, 3, f"=Val!B{VR['tpdcf_s']}").number_format = "#,##0.00"; RC.cell(r, 4, f"=Val!B{VR['tpdcf']}").number_format = "#,##0.00"; RC.cell(r, 5, f"=Val!B{VR['px']}").number_format = "#,##0.00"; RC.cell(r, 7, "欄位：加回目標價／主值 DCF 目標價／現價"); r += 1
RCR["rc_txt"] = r; RC.cell(r, 1, f'="DCF（股票薪酬加回）每股價值 "&TEXT(C{RCR["rc_psv_s"]},"#,##0.00")&" 美元、目標價 "&TEXT(C{RCR["rc_tp_s"]},"#,##0.00")&" 美元，為現價 "&TEXT(E{RCR["rc_tp_s"]},"#,##0.00")&" 美元的 "&TEXT(C{RCR["rc_tp_s"]}/E{RCR["rc_tp_s"]},"0%")&"；主值 DCF 目標價 "&TEXT(D{RCR["rc_tp_s"]},"#,##0.00")&" 美元把股票薪酬視為費用。此列只作對照，不進加權目標價與評等。"'); r += 2
RC.cell(r, 1, "8. 現金稅與模型稅率（原則第 40 條）").font = BOLD; r += 1
RCR["cash_tax"] = r; RC.cell(r, 1, "FY2026 現金稅（公司預期）vs 模型所得稅費用（情境 B，US$M）"); c = RC.cell(r, 3, fv("tax.cash.fy26g")); c.font = BLUE; RC.cell(r, 4, f"=PL!H{PB_['tax']}").number_format = "#,##0.0"; RC.cell(r, 7, "現金稅遠低於帳面稅（美國遞延所得稅資產全額備抵、股票薪酬扣除）；模型以帳面稅率，DCF 偏保守"); r += 2
RC.cell(r, 1, "9. 未找到").font = BOLD; r += 1
for t_ in ("未找到：公司不揭露 AI 產品（Bits AI 等）營收、AI 推論成本金額與 token 量；每任務 token 由定價頁與法說「credits 價格貼近推論成本」校準。",
           "未找到：最大客戶身分與占營收比（公司未揭露；10-Q 只稱其降低用量，Q4 指引以最低承諾計）。",
           "未找到：FY2027 公司指引；財報後 FY2027 營收與 EPS 共識的金額（只有成長 23% 與財報前值）。",
           "未找到：RSU 預期沒收率；對照列 Forward P/E 的分母期間與租賃處理；賣方目標價家數。同業 P/S 中 MDB、DT、ESTC 價格日為 2026-11-06，逾期（Val 0.6 標示）。"):
    RC.cell(r, 1, t_); r += 1
RC.column_dimensions["A"].width = 80
for c_ in "BCDEF": RC.column_dimensions[c_].width = 16
RC.column_dimensions["G"].width = 80

# ====================================================================================================== Sens
SE["A1"] = "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算；加權目標價與 DCF 目標價取自重算後的 Val 頁（sens_datadog.py；建置時寫入數值快照）。"
SE["A2"] = ("相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。token 價格高案只動成本一邊，標「上界」；另列價格與成本聯動、模型路由兩個組合案例（G2 第 2 條）。"
            "股票薪酬與費用率的「全期持平」案為歷史斜率延伸的下緣（G3 第 1 條）。cons_tp 只作對照，不進模型，不列。tk_tokens 為推得值，由 cost_sh、cr_task、cr_px 驅動，不另列。")
hdr(SE, 3, ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "收支關卡 B（五年，US$M）", "DCF 目標價", "投資建議"])
SENS_DEF = []
for row_ in ROWS:
    if row_[0] == "sec": continue
    key, item, unit, lo, base, hi = row_[:6]
    if key in ("cons_tp",): continue
    for side, v in (("低", lo), ("高", hi)):
        if v is None: continue
        if isinstance(base, (int, float)) and not isinstance(base, bool) and v == base: continue
        if isinstance(v, (datetime.date,)): continue
        SENS_DEF.append((key, side, v, item))
COMBO = {"@link": ("tk_px、pt_k", "tk_px=%r;pt_k=1.0" % tkv("IF_PriceRef_Astra"), "價格與成本聯動：token 價格取 Astra 層，同時 credit 價格完全轉嫁（pt_k＝1）"),
         "@route": ("tk_px、route_sh", "tk_px=%r;route_sh=0.7" % tkv("IF_PriceRef_Astra"), "模型組合：高價 Astra 層，但 70% 任務路由到 Terra 層便宜模型")}
COMBO["@ps_only"] = ("rel_m", "rel_m=%r" % _R0["ps_reading"], "相對倍數只用 PS 讀數（近 20 財季中位）")
COMBO["@psg_only"] = ("rel_m", "rel_m=%r" % _R0["psg_reading"], "相對倍數只用 PSG 讀數（成長調整後）")
for ck, (kt, vv_, lab_) in COMBO.items(): SENS_DEF.append((ck, "聯動", vv_, lab_))
r = 4
SE_ROWS = []
sj = SENSJ.get("cases", {})
for k, side, v, lab in SENS_DEF:
    SE.cell(r, 1, lab if k.startswith("@") else (f"{lab}：{side}（上界：只動成本一邊，未連動價格與路由）" if (k == "tk_px" and side == "高") else f"{lab}：{side}"))
    SE.cell(r, 2, f"基準 → {v if not isinstance(v, float) else round(v, 4)}"); SE.cell(r, 3, side)
    cs = sj.get(f"{k}|{side}")
    if cs:
        for cc, kk, fm in ((4, "tp", "#,##0.00"), (8, "gate", "#,##0.0"), (9, "dcf", "#,##0.00")):
            c = SE.cell(r, cc, cs[kk]); c.number_format = fm; c.font = BLUE
        SE.cell(r, 10, cs["rec"])
    SE.cell(r, 5, f'=IF(D{r}="","",D{r}/$D$BASEROW-1)').number_format = "+0.0%;-0.0%;0.0%"
    SE.cell(r, 6, COMBO[k][0] if k.startswith("@") else k); SE.cell(r, 7, "重算值（sens_datadog.py）" if cs else "未重算")
    SE_ROWS.append(r); r += 1
SE_END = r - 1
r += 1
SE_BASE = r; SE.cell(r, 1, "基準加權目標價（Val 頁）"); SE.cell(r, 4, f"=Val!B{VR['tp']}").number_format = "#,##0.00"
for rr_ in SE_ROWS: SE[f"E{rr_}"].value = SE[f"E{rr_}"].value.replace("BASEROW", str(SE_BASE))
r += 1
SE_STALE = r; SE.cell(r, 1, "Sens 重跑時的基準加權目標價"); c = SE.cell(r, 4, SENSJ.get("base_tp")); c.number_format = "#,##0.00"; c.font = BLUE; r += 1
FP_RNG = f"In!$D$4:$F${IN_LAST}"
SE_FP = r; SE.cell(r, 1, "輸入指紋（本檔即時重算）"); SE.cell(r, 4, f"=SUMPRODUCT({FP_RNG},ROW({FP_RNG})*3+COLUMN({FP_RNG}))").number_format = "0.000000"; r += 1
SE_FP0 = r; SE.cell(r, 1, "輸入指紋（Sens 重跑時的基準檔）"); c = SE.cell(r, 4, SENSJ.get("fp")); c.number_format = "0.000000"; c.font = BLUE; r += 1
SE_STATUS = r; SE.cell(r, 1, "模型狀態（CURRENT＝Sens、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 SENS=1 bash make_datadog.sh）").font = BOLD
SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{SE_FP0}),ISNUMBER(D{SE_STALE}),ABS(D{SE_FP}-D{SE_FP0})<=0.000000001*MAX(1,ABS(D{SE_FP0})),ABS(D{SE_STALE}-D{SE_BASE})<0.005),"CURRENT","STALE")'); r += 1
SE_BID = r; SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, BUILD_ID); r += 1
SE_TORN = r; SE.cell(r, 1, "龍捲風圖說明（HTML）：|相對基準| ≥ 10% 的案例列在前，未達者摺疊；全部案例見上表。"); r += 1
SE.column_dimensions["A"].width = 70; SE.column_dimensions["B"].width = 26; SE.column_dimensions["G"].width = 30

# ====================================================================================================== Sources
SO["A1"] = "Sources — 來源與版本"
hdr(SO, 3, ["編號", "文件", "日期", "網址", "標記", "用途"])
SRCS = [("事實總帳 data/facts_datadog.csv（facts_datadog.py；每筆含出處、擷取日、標記、來源誘因）", "2026-11-16", "repo 內", "各筆標記", "Hist、Q、In 頁全部實績與多數輸入"),
        ("Datadog 2026Q3 財報新聞稿與 10-Q（SEC EDGAR）", "2026-11-04", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001561550", "Verified", "營收、損益、資產負債、RPO、股數、最大客戶描述"),
        ("Datadog 投資人關係：財報新聞稿與法說逐字稿", "2026-11-04", "https://investors.datadoghq.com/", "Interested-party", "指引、AI 原生客戶占比、非 AI 成長、credits 定價評論"),
        ("Datadog AI Credits 定價頁", "2026-11-16", "https://www.datadoghq.com/pricing/", "Verified", "年約每 credit 1.00、隨需 1.30、Bits Investigation 約 4 credits（官網估算，Interested-party）"),
        (f"Tokenomics 快照 data/tokenomics_snapshot_v22.0.json（{TK['source']['file']}，commit {TK['source']['commit'][:7]}）", TK["source"]["extractedAt"], "repo 內（YinchenChang/Tokenomics）", "Analogy", "IF_PriceRef_Luna／Terra／Sol／Astra、IF_TaskTokMult"),
        ("stockanalysis.com 統計頁（同業與 Datadog Forward 倍數、β）", "2026-11-06／11-13", "https://stockanalysis.com/", "Interested-party", "對照列 Forward P/E 與本公司現行倍數（SNOW、NET 11-13；MDB、DT、ESTC 11-06＝逾期）"),
        ("TipRanks／TIKR 共識", "2026-11-12", "https://www.tipranks.com/stocks/ddog/forecast", "Interested-party", "平均目標價、FY2027 成長共識"),
        ("10 年期美國公債殖利率與 DDOG 收盤價（兩來源）", "2026-11-13", "見事實總帳 mkt.* 出處欄", "Verified", "rf、現價")]
for i, (doc, dt, url, tag, use) in enumerate(SRCS, 1):
    for j, v in enumerate((f"S{i:02d}", doc, dt, url, tag, use), 1): SO.cell(3 + i, j, v)
for col, w in zip("ABCDEF", [6, 80, 18, 60, 18, 60]): SO.column_dimensions[col].width = w

# ====================================================================================================== README
VB = lambda k, c="B": f"Val!{c}{VR[k]}"
TB = T["B"]
RD["A1"] = (f'="Datadog（DDOG）個股模型 {VER}｜build_id "&Sens!D{SE_BID}&"｜資料截止日 "&TEXT(data_date,"yyyy-mm-dd")&"｜成品日 {VDATE}｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'
            f'&"｜評價日 "&TEXT(val_date,"yyyy-mm-dd")&"｜目標時點 "&TEXT(tp_date,"yyyy-mm-dd")&"（Tokenomics 快照 v22.0；命題＝原型 U 按用量計價，單位經濟在 In 頁自建，v0.2 改讀 unit_economics）"')
RD["A2"] = ("v0.2.2 要點（G8：P/S 主腿股數改評價日稀釋股數、並列扣非控制權益＝0；v0.2.1：同業 P/S 改現價日 peer_ps_20261113，G7；相對倍數改讀歷史相對倍數 geo_base、無風險利率統一為 5.10%、並列股票薪酬加回 DCF 對照）：命題＋官方單一分部 → 損益表；套用共通必修 G2 十條與 G3 四條；DCF 為 FCFF（股票薪酬視為費用）到 FY2032，Comparable 主腿為同業 P/S（與 PS／PSG 相對值同口徑，G6），Forward P/E 只列對照；"
            "股票薪酬占營收依歷史規模經濟斜率延伸（下限 sbc_floor），其餘比率持平；三情境只在 AI 驅動分化（非 AI 與 AI 原生客戶用量、AI 原生客戶價格讓利、AI credits）。")
RD["A3"] = "命題：Datadog 的營收按用量計價（營收＝客戶用量 × 單價）。AI 原生客戶帶來用量爆發，但也以大額承諾換取較低單價；AI 產品 credits 帶來新營收卻需向模型商買推論。本模型檢驗 AI 淨效果（增量 − 價格讓利 − 推論成本）的正負，以及它是否足以支持現價。"
RD["A4"] = (f'="核心問題：FY2032E 的 AI 淨效果（營業利益前毛利）在情境 A／B／C 分別為 "&TEXT({VB("ai_net")},"#,##0")&"／"&TEXT({VB("ai_net","C")},"#,##0")&"／"&TEXT({VB("ai_net","D")},"#,##0")&" 百萬美元；基準下 AI 增量是否大於價格讓利與推論成本。"')
RD["A5"] = (f'="投資建議："&{VB("rec")}&"——加權目標價 "&TEXT({VB("tp")},"#,##0.00")&" 美元（DCF "&TEXT({VB("tpdcf")},"#,##0.00")&" 與 Comparable "&TEXT({VB("tpcomp")},"#,##0.00")&" 各 "&TEXT(w_dcf,"0%")&"／"&TEXT(1-w_dcf,"0%")&"），上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"（現價 "&TEXT(px,"#,##0.00")&"，"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）；收支關卡："&{VB("gate_chk")}&"。"')
RD["A6"] = (f'="重點觀察 1：兩個方法差很大。DCF 把股票薪酬當成本，只有 "&TEXT({VB("tpdcf")},"#,##0.00")&" 美元；Comparable 用同業 P/S（乘歷史相對倍數），達 "&TEXT({VB("tpcomp")},"#,##0.00")&" 美元；落差 "&TEXT({VB("g_tot")},"+#,##0.00")&" 美元，其中股票薪酬佔 "&TEXT({VB("g_sbc")},"+#,##0.00")&"，其餘是倍數與成長。"')
RD["A7"] = (f'="重點觀察 2：現價隱含 FY2028–FY2032 GAAP 營業利益率約 "&TEXT(RevDCF!B{RVR["m1"]},"0.0%")&"，模型情境 B 平均 "&TEXT(RevDCF!B{RVR["m0"]},"0.0%")&"；現行 Forward P/S 為同業中位數的 "&TEXT({VB("c_ps","D")},"0.00")&" 倍。"')
RD["A8"] = (f'="重點觀察 3：FY2032E AI 淨效果 A／B／C＝"&TEXT({VB("ai_net")},"#,##0")&"／"&TEXT({VB("ai_net","C")},"#,##0")&"／"&TEXT({VB("ai_net","D")},"#,##0")&" 百萬美元；情境 B 推論成本 FY2032E "&TEXT(Thesis!N{TB["inf"]},"#,##0")&"，AI credits 轉售毛利率 "&TEXT(Thesis!N{TB["rm"]},"0%")&"。"')
RD["A9"] = (f'="重點觀察 4：FY2026E 營收模型 "&TEXT(Recon!C{RCR["fyrev"]},"#,##0")&" 對指引中點 "&TEXT(Recon!F{RCR["fyrev"]},"+0.0%;-0.0%")&"（公司近三季皆超出指引上緣）；FY2027E 成長情境 B "&TEXT(Recon!D{RCR["g27"]},"0.0%")&"，共識 "&TEXT(Recon!F{RCR["g27"]},"0.0%")&"。"')
RD["A10"] = (f'="重點觀察 5：收支關卡（FY2026–FY2030 自由現金流）A／B／C＝"&TEXT({VB("gate_A")},"#,##0")&"／"&TEXT({VB("fcf_gate")},"#,##0")&"／"&TEXT({VB("gate_C")},"#,##0")&" 百萬美元；WACC "&TEXT({VB("wacc")},"0.00%")&"，終值占 EV "&TEXT({VB("tvsh","C")},"0%")&"。"')
RD["A11"] = (f'="橫幅："&{VB("rec")}&"；加權目標價 "&TEXT({VB("tp")},"#,##0")&" 美元，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"。DCF 因股票薪酬視為成本而遠低於現價，Comparable 因同業倍數而偏高，兩法差距大，結論取決於如何看待股票薪酬與倍數。"')
RD["A13"] = (f'="情境分化依據：A／B／C 只在四個 AI 驅動分化——非 AI 客戶 FY2027 用量成長（"&TEXT(vn_27_A,"0%")&"／"&TEXT(vn_27_B,"0%")&"／"&TEXT(vn_27_C,"0%")&"，依近五季非 AI 成長 high 20s 的延續或減速）、AI 原生客戶用量成長（"&TEXT(va_27_A,"0%")&"／"&TEXT(va_27_B,"0%")&"／"&TEXT(va_27_C,"0%")&"，依代理工作負載放大 vs 最大客戶停留最低承諾）、AI 原生客戶單價變動（"&TEXT(pa_A,"0%")&"／"&TEXT(pa_B,"0%")&"／"&TEXT(pa_C,"0%")&"，依提前續約讓利）、AI credits FY2027 營收（"&TEXT(ai_27_A,"#,##0")&"／"&TEXT(ai_27_B,"#,##0")&"／"&TEXT(ai_27_C,"#,##0")&" 百萬美元，公司未揭露，Assumed）。"')
RD["A14"] = ("與標準規格的偏離：(1) 2029 可轉換票據預設視為債務（現價與基準每股價值皆低於轉換價約 217.60），conv_if＝1 為 Sens；(2) 費用率基準：股票薪酬占營收依歷史斜率延伸（FY2023 至近四季），non-GAAP 費用率三項合計歷史無一致下降→持平；"
             "(3) Comparable 主腿為同業 P/S（2026-11-13 現價，MDB、DT、ESTC 為 2026-11-06，逾期），乘相對倍數 geo_base；盈餘型倍數只作對照；(4) RSU 預期沒收率未找到，基準 0。")
RD["A15"] = ("不適用：Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照——Datadog 不自建 GPU 算力，只向模型商購買推論，推論成本讀 Tokenomics 快照的市場價名稱。"
             "不適用：三層協調（Coord 頁、C0、兩讀數、Siliconomics）——非 Siliconomics 覆蓋；對帳中兩讀數目標價亦不適用。")
RD["A16"] = ("已知限制（未找到）：AI 產品營收與推論成本金額、最大客戶身分與占比、FY2027 公司指引、財報後 FY2027 營收與 EPS 共識金額、RSU 預期沒收率、同業倍數分母期間與租賃處理、賣方目標價家數；"
             "MDB、DT、ESTC 倍數為 2026-11-06（逾期）。詳見 data/research_datadog.md 第九節。")
RD["A17"] = ("共通必修 G2 十條：(1) 每任務 token 由公司揭露校準（In 頁 tk_tokens），聊天基數類推只列 Recon 第 5 節對照，token 價格層級標 Analogy；(2) token 價格高案標上界，另列聯動與路由組合案；(3) 比率基準持平於近四季，歷史有證據者（股票薪酬）延伸，收斂只放 Sens；"
             "(4) 股數取封面、含 RSU 與選擇權庫藏股法、票據 if-converted 為 Sens；(5) 淨債務只計票據本金與融資租賃，營業租賃不計；(6) Comparable 倍數標分母期間、同業標準，並揭露現行倍數 ÷ 同業中位數；"
             "(7) Val 頁有估值評等、入不敷出判定用、投資建議，check_rating_gate 已跑；(8) README 文字數字由公式組字，落差以公式分解；(9) 主導變數整本重算反解在 Recon 第 6 節；(10) rf、同業倍數逾期標示在 Val 0.6。")
RD["A18"] = f'="合理性關卡（G3 第 3 條）：觸發 "&TEXT(Recon!B{RCR["gt_n"]},"0")&" 項。診斷結論："&Recon!B{RCR["diag"]}'
RD["A19"] = "驅動 → 推導對照表（每頁輸入、推導量、因果方向）"
hdr(RD, 20, ["頁", "輸入", "推導量", "因果方向"])
MAP = [("In→Thesis", "非 AI 與 AI 原生客戶用量成長、單價變動、AI credits", "三條營收子列（用量 × 單價）", "用量↑→營收↑；單價讓利↑→營收↓"),
       ("In→Thesis", "每任務 token、token 價格（Tokenomics）、credit 價格轉嫁", "推論成本與 AI credits 轉售毛利率", "token 價格↑→推論成本↑，壓低毛利"),
       ("Thesis→PL", "營收、用量、推論成本", "毛利、non-GAAP 費用、GAAP 營業利益", "單價讓利壓低毛利率（成本隨用量）"),
       ("PL→CF", "淨利、折舊、股票薪酬、營運資金、資本支出", "自由現金流與收支關卡", "遞延收入增加為流入"),
       ("PL／CF→DCF", "GAAP 營業利益、稅、折舊、資本支出、營運資金", "FCFF 與現值", "股票薪酬視為費用"),
       ("DCF／PL→Val", "WACC、永續成長、同業 P/S（relmult_hist）", "DCF、Comparable、加權目標價、評等", "兩法依權重加權；評等依上檔空間與收支關卡"),
       ("Val→Sens", "每個有區間的輸入", "加權目標價相對基準", "一次改一個輸入，整本重算")]
for i_, rowv in enumerate(MAP):
    for j, v in enumerate(rowv): RD.cell(21 + i_, j + 1, v)
RD["A29"] = "市場資料日期、失效條件與風險登錄"
RD["A30"] = (f'="市場資料日期：現價 "&TEXT(px_date,"yyyy-mm-dd")&" 收盤（兩來源前收價相符）；10 年期美國公債 "&TEXT(rf,"0.00%")&"（"&TEXT(rf_date,"yyyy-mm-dd")&"，"&Val!C{VR["o_rf"]}&"）；共識目標價 "&TEXT(cons_date,"yyyy-mm-dd")&"（"&Val!C{VR["o_cons"]}&"）；同業倍數資料日 "&TEXT(MIN(pm_dp_snow,pm_dp_net,pm_dp_mdb,pm_dp_dt,pm_dp_estc),"yyyy-mm-dd")&" 至 "&TEXT(MAX(pm_dt_snow,pm_dt_net,pm_dt_mdb,pm_dt_dt,pm_dt_estc),"yyyy-mm-dd")&"（MDB、DT、ESTC 逾期）。"')
RD["A31"] = "主要失效條件"
RD["A32"] = (f'="(1) 最大客戶續約後用量低於 Q4 指引假設的最低承諾，或同群體第二、第三大客戶跟進：情境 C 的 AI 原生客戶用量成長（"&TEXT(va_27_C,"0%")&"）需再下修。(2) 公司揭露 AI 產品營收或 token 用量後，與 Assumed 區間差距大。(3) 股票薪酬占營收下降速度低於歷史斜率：加權目標價走向 Sens 低案。(4) 同業倍數逾期列更新後，Comparable 重算。"')
RD["A33"] = "風險登錄"
hdr(RD, 34, ["風險", "是否建模", "理由", "觀察指標"])
RISKS = [("最大 AI 原生客戶降低用量或自建", "是（情境 C、事件型 lc_cut）", "10-Q 稱最大客戶 2026Q3 降低用量，Q4 指引以最低承諾計", "季報客戶集中度、AI 原生占營收比"),
         ("AI 原生客戶提前續約讓利", "是（單價變動）", "以較低單價提前續約已發生，幅度未揭露", "非 AI 與 AI 原生營收成長差"),
         ("推論成本與模型價格", "是", "讀 Tokenomics 市場價；高案 Astra 層（上界）", "AI credits 營收與毛利率"),
         ("併購現金支出", "否", "9M26 約 115.6，不可預測；列觀察", "現金與有價證券"),
         ("內建可觀察性與雲端原生替代（雲端大廠自有監控）", "否", "無可量化輸入；反映於非 AI 用量成長的區間", "新客戶數、淨營收留存率")]
for i_, rowv in enumerate(RISKS):
    for j, v in enumerate(rowv): RD.cell(35 + i_, j + 1, v)
RD["A41"] = "工作表：README／In／Hist／Q／Thesis／Seg／PL／CF／DCF／Val／RevDCF／Recon／Sens／Sources。成品 dist/<日期>_Datadog個股模型_v0.2.1.xlsx 與 .html。"
RD["A42"] = (f'="DCF 與 Comparable 落差歸因（情境 B）：Comparable "&TEXT({VB("comp","C")},"#,##0.00")&" − DCF "&TEXT({VB("dcf12","C")},"#,##0.00")&" ＝ "&TEXT({VB("g_tot")},"+#,##0.00;-#,##0.00")&" 美元／股：股票薪酬 "&TEXT({VB("g_sbc")},"+#,##0.00;-#,##0.00")&"、倍數與成長殘差 "&TEXT({VB("g_mult")},"+#,##0.00;-#,##0.00")&"；槓桿 兩法皆為股權每股價值（每股淨現金 "&TEXT({VB("g_lev")},"0.00")&" 已含）；時點 兩法皆在目標時點（DCF 評價日起以 ke 滾動 "&TEXT({VB("g_time")},"+0.00")&" 已含在 DCF 內）。"')
RD["A43"] = "無預設觀點：本版未預先設定任何投資觀點，結論全由 Assumed 輸入與公開資料推出。"
for rr_ in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 30, 32, 42, 43):
    RD[f"A{rr_}"].alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
RD.column_dimensions["A"].width = 120
for c_ in "BCD": RD.column_dimensions[c_].width = 36

# ====================================================================================================== Seg
SG["A1"] = "Seg — 官方營收分部：Datadog 為單一應報導部門（10-K Note 2：Datadog 平台）；本模型的三條子列為命題拆分（非官方分部）。"
SG["A2"] = "官方分部：Datadog 平台（單一應報導部門）＝合併營收。子列：非 AI 原生客戶平台用量、AI 原生客戶平台用量、AI 產品 credits 營收（Thesis 頁推導，加總＝分部）。"
hdr(SG, 4, ["項目", "FY2025A", "說明"])
SG["A5"] = SEGL; SG["B5"] = "=h_rev_25"; SG["B5"].number_format = "#,##0.0"; SG["C5"] = "10-K：單一應報導部門；營收 3,427.158"
SG.column_dimensions["A"].width = 60; SG.column_dimensions["C"].width = 70

# PL 頁圖表資料（連續欄：FY2023A–FY2025A、FY2026E–FY2032E；供 HTML 折線與長條圖）
r = PL.max_row + 3
PL.cell(r, 1, "圖表資料（情境 B；連續欄供 HTML 圖表使用，值同上方標準損益表）").font = BOLD; r += 1
CHR = {"hdr": r}
PL.cell(r, 1, "項目"); PL.cell(r, 2, "單位")
SRCC = ["C", "D", "E", "H", "I", "J", "K", "L", "M", "N"]; DSTC = "CDEFGHIJKL"
for d_, s_ in zip(DSTC, SRCC): PL[f"{d_}{r}"] = f"={s_}{PL_HDR}"
r += 1
for k, lab, unit, src, fm in (("rev", "營收", "US$M", "rev", "#,##0.0"), ("gm", "毛利率", "%", "gm", "0.0%"), ("om", "營業利益率", "%", "om", "0.0%"), ("nm", "淨利率", "%", "nm", "0.0%")):
    CHR[k] = r; PL.cell(r, 1, lab); PL.cell(r, 2, unit)
    for d_, s_ in zip(DSTC, SRCC): PL[f"{d_}{r}"] = f"={s_}{PB_[src]}"; PL[f"{d_}{r}"].number_format = fm
    r += 1
ROWMAP = {"IN": IN_ROW, "PR": P_, "TR": T, "CR": C_, "DR": DR_, "IX": IS, "VR": VR, "RVR": RVR, "RCR": RCR, "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE, "SE_STALE": SE_STALE, "SE_STATUS": SE_STATUS,
          "SE_BID": SE_BID, "SE_END": SE_END, "SE_TORN": SE_TORN, "SE_FP": SE_FP, "PL_HDR": PL_HDR, "PL_NOTE": PL_NOTE, "PL_SUM": PL_SUM, "PL_SUM2": PL_SUM2, "PL_SUM3": PL_SUM3, "CHR": CHR, "RD_ROWS": {"hero": 11}, "GATE": GATE, "SCN": SCN, "SCN_H": SCN_H, "BUILD_ID": BUILD_ID,
          "SENS_DEF": [(k, s_, str(v)) for k, s_, v, _ in SENS_DEF], "IN_LAST": IN_LAST, "Q_CHK": [Q_CHK1, Q_CHK2, Q_CHK3]}
if os.environ.get("ROWMAP"):
    json.dump(ROWMAP, open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
wb.save(A.out)
print("saved", A.out, BUILD_ID)

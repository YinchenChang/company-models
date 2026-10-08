"""Anthropic v0.1-A3：算力頁 Compute 與成本頁 Cost（工作單 A3；規格 D6 r1、D7–D13、D10 r1、D21、D22）。

結構沿用 OpenAI v0.6 builder/p3.py、p4.py（同名具名範圍：CMP_*、COST_*），差異：
- 加速器族 × 世代（D6、D6 r1）：NVIDIA Hopper／GB200／GB300／VR200 直接取 TK IF_TokGW_*；TPU v6e、Trainium2＝Hopper × NonNV 產出比；
  TPU v7（含次世代 TPU）、Trainium3、AMD MI455X＝VR200 × NonNV 產出比。合約價與供應商持有成本 × NonNV 持有比（D9）。
- 供給 GW 只由算力合約加總（D10 r1）；機房租約（TeraWulf、Riot、Hut 8、Nexus、Cipher）只列對照。逐合約依揭露情形分三型：
  金額型（只揭露金額與期間）：GW＝年付款 ÷（合約價 × 持有比）；GW 型（只揭露 GW，或金額只有下限、期間不明）：付款＝GW × 合約價 × 持有比；
  雙揭露型（金額＋期間＋GW 皆有）：供給＝揭露 GW、付款＝金額時程。另有既有雲端（2025 前簽；2025＝說明書實際算力支出）與 SpaceX 月費。
- 自建（Fluidstack $50B；D11）：資本支出均攤，投產落後 Inputs 年數後計入供給（÷ TK IF_CapexTotal）。
- η（D8）＝2025 token 換算推論 GW ÷ 2025 推論支出（說明書算力 7.33 − 訓練 4.1）換算 GW。容量上限 2025、2026 固定 1（Inputs cap_fixed_year）。
Python 只產生結構；數值一律引用 SRC_ANT／Inputs／TK_／DEM_／REV_ 具名範圍或同簿儲存格（E6）。
"""
from __future__ import annotations

from a2 import FMT, TIERS, YC, YEARS, Sheet
from common import put

FMT.update({"GW": "0.000", "$B/GW/年": "0.00", "M tok/GW/年": "#,##0", "旗標": "0", "人": "#,##0", "$M/人/年": "0.000", "年": "0", "月": "0", "M tok/年": "#,##0"})

GENS = ("hopper", "gb200", "gb300", "vr200")                     # NVIDIA 世代（TK 世代表頭前 4 個）
GEN_ZH = {"hopper": "Hopper", "gb200": "GB200", "gb300": "GB300", "vr200": "VR200"}
TK_TIER = {"top": "Astra", "mid": "Sol", "low": "Luna"}
# 加速器族：key、中文、基準 NVIDIA 世代、NonNV 族（None＝NVIDIA）、具名範圍後綴
FAMS = [
    ("hopper", "NVIDIA Hopper", "hopper", None, "Hopper"),
    ("gb200", "NVIDIA GB200", "gb200", None, "GB200"),
    ("gb300", "NVIDIA GB300", "gb300", None, "GB300"),
    ("vr200", "NVIDIA VR200（Vera Rubin）", "vr200", None, "VR200"),
    ("tpu6", "Google TPU v6e（＝Hopper × TPU 產出比）", "hopper", "TPUv7", "TPUv6e"),
    ("tpu7", "Google TPU v7 Ironwood／次世代 TPU（＝VR200 × TPU 產出比）", "vr200", "TPUv7", "TPUv7"),
    ("trn2", "AWS Trainium2（＝Hopper × Trainium 產出比）", "hopper", "Trn3", "Trn2"),
    ("trn3", "AWS Trainium3（＝VR200 × Trainium 產出比）", "vr200", "Trn3", "Trn3"),
    ("mi455", "AMD MI455X（Helios；＝VR200 × AMD 產出比）", "vr200", "MI455X", "MI455X"),
]
FAM = {f[0]: f for f in FAMS}
NNV = ("TPUv7", "Trn3", "MI455X")
# 既有雲端（2025 前簽）組合：Inputs 設 Trainium2、Hopper，TPU v6e＝餘額
LEGACY = (("trn2", "legacy_share_trn2"), ("hopper", "legacy_share_hopper"), ("tpu6", None))

# 逐合約（D10 r1）：key、名稱、族、型、參數、招股書六家合作方（$518B 對帳）、說明
#   amt：金額型（T 總額、s 起、e 迄）；both：雙揭露型（另 G 揭露 GW 公式）；gw：GW 型（G、s、term、upto）；monthly：月費（SpaceX）
CONTRACTS = [
    ("google", "Google Cloud TPU（2025-10；招股書 Google ≥$111.1B，2026-04～2033-07）", "tpu7", "amt",
     dict(T="{SRC:cmp_prospectus_google}_Lo", s="{INP:google_start}", e="{INP:google_end}"), "Google",
     "金額型：招股書 Google 承諾；揭露的「2026 年上線遠超過 1 GW」與「最多 100 萬顆 TPU」只作對照（以承諾金額為準）"),
    ("broadcom", "Google／Broadcom 次世代 TPU（2026-04；8-K 約 3.5 GW；五年 TPU 租賃 $125.2B）", "tpu7", "both",
     dict(T="{SRC:cmp_broadcom_tpu_lease_125b}", s="{INP:broadcom_start}", e="{INP:broadcom_end}",
          G="({SRC:cmp_google_broadcom_2026_gw_b}+{INP:broadcom_gw_switch}*({SRC:cmp_google_broadcom_2026_gw_c}-{SRC:cmp_google_broadcom_2026_gw_b}))"), "Broadcom",
     "雙揭露型：供給＝3.5 GW（Broadcom 8-K，Verified；公司自述 5 GW 為區間）；付款＝五年 TPU 租賃 $125.2B 時程（招股書 Broadcom 設備租賃 $161.2B 的其餘 $36B 未計，見報告）"),
    ("aws", "Amazon 擴大協議（2026-04；招股書 $110B，2026-05～2036-04；最多 5 GW）", "trn3", "amt",
     dict(T="{SRC:cmp_prospectus_amazon}", s="{INP:aws_start}", e="{INP:aws_end}"), "Amazon",
     "金額型：「最多 5 GW」為上限、只作對照；晶片 Trainium2–4，本模型歸 Trainium3（Decision）"),
    ("azure", "Microsoft Azure＋NVIDIA（2025-11；招股書 $31.4B，2026-11～2033-05；最多 1 GW）", "gb300", "amt",
     dict(T="{SRC:cmp_msft_azure_prospectus}", s="{INP:azure_start}", e="{INP:azure_end}"), "Microsoft",
     "金額型：Grace Blackwell 與 Vera Rubin，本模型歸 GB300（Decision）；「最多 1 GW」只作對照"),
    ("amd", "AMD MI455X（2026-07；最多 2 GW；招股書 >$20B）", "mi455", "gw",
     dict(G="{SRC:cmp_amd_gw}_Hi*{INP:upto_share}", s="{INP:amd_start}", term="{INP:amd_term}"), "AMD",
     "GW 型：金額只有下限（>$20B）且期間未揭露 → 付款＝GW × 合約價 × AMD 持有比；「最多 2 GW」× 計入比例（工作單預設 0.8）"),
    ("spacex", "SpaceX／xAI Colossus（2026-05；$1.25B／月至 2029-05，SpaceX S-1）", "hopper", "monthly",
     dict(F="{SRC:cmp_spacex_monthly_fee}", s="{INP:spacex_start}", e="{INP:spacex_end}", m1="{INP:spacex_m1}", m2="{INP:spacex_m2}"), "xAI／SpaceX",
     "金額型（月費，Verified）：GW＝付款 ÷ 合約價；Colossus 1「>300 MW、>22 萬顆 GPU」只作對照（月費應涵蓋擴大後容量）；招股書上限 $84.5B 為區間。晶片以 H100／H200 為主 → Hopper（Decision）"),
    ("nscale", "Nscale Monarch（2026-08；$45B／六年；460 MW；Vera Rubin）", "vr200", "both",
     dict(T="{SRC:cmp_nscale_usd}", s="{INP:nscale_start}", e="{INP:nscale_end}", G="{SRC:cmp_nscale_mw}/{INP:mw_to_gw}"), None,
     "雙揭露型：供給＝460 MW；付款＝$45B 時程"),
    ("lambda", "Lambda（2026-09；$35B，多年；350 MW；NVIDIA）", "gb300", "gw",
     dict(G="{SRC:cmp_lambda_mw}/{INP:mw_to_gw}", s="{INP:lambda_start}", term="{INP:lambda_term}"), None,
     "GW 型：期間未揭露 → 付款＝GW × 合約價；晶片未揭露，歸 GB300（Decision）；隱含總額對 $35B 見對帳列"),
    ("volta", "Volta Infra（2026-08；$10B／六年；133 MW；Vera Rubin）", "vr200", "both",
     dict(T="{SRC:cmp_volta_usd}", s="{INP:volta_start}", e="{INP:volta_end}", G="{SRC:cmp_volta_mw}/{INP:mw_to_gw}"), None,
     "雙揭露型：供給＝133 MW；付款＝$10B 時程"),
    ("akamai", "Akamai（2026-05；$1.8B／七年；RTX PRO 6000 Blackwell）", "gb200", "amt",
     dict(T="{SRC:cmp_akamai_usd}", s="{INP:akamai_start}", e="{INP:akamai_end}"), None,
     "金額型：容量未揭露；Blackwell 推論卡以 GB200 代理（Decision）"),
]
CT = {c[0]: c for c in CONTRACTS}
LEASES = [  # 機房租約（只含建物與電力；D10 r1：不另計 GW）：key、名稱、MW 公式、說明
    ("terawulf", "TeraWulf Hawesville（KY；IT 401 MW；20 年約 $19B；直接承租）", "{SRC:cmp_terawulf_ky_mw}", "硬體為 Google TPU（第三方）→ 容量已含於 TPU 合約"),
    ("riot", "Riot Rockdale（TX；191 MW；20 年 $9.1B）", "{SRC:cmp_riot_rockdale_mw}", "硬體為 AMD MI450（第三方）→ 容量已含於 AMD 合約"),
    ("hut8", "Hut 8 River Bend（LA；245 MW；經 Fluidstack，15 年 $7B）", "{SRC:cmp_hut8_riverbend_mw}", "經 Fluidstack；硬體 Google TPU → 已含於 TPU 合約／自建"),
    ("nexus", "Nexus Hubbard（TX；首期 500 MW；Google 擔保）", "{SRC:cmp_nexus_hubbard_mw}", "Google 擔保 → 容量已含於 TPU 合約"),
    ("cipher", "Cipher Barber Lake（TX；IT 207 MW；Epoch 歸屬，Analogy）", "{SRC:cmp_barber_lake_epoch}", "經 Fluidstack；TPU v7 → 已含於 TPU 合約／自建"),
]
PARTNERS = [("Google", "cmp_prospectus_google", "_Lo"), ("Amazon", "cmp_prospectus_amazon", ""), ("Microsoft", "cmp_msft_azure_prospectus", ""),
            ("Broadcom", "cmp_prospectus_broadcom", ""), ("xAI／SpaceX", "cmp_spacex_total_prospectus", "_Hi"), ("AMD", "cmp_amd_usd_prospectus", "_Lo")]


def build(ctx):
    D, wb, snap = ctx.D, ctx.wb, ctx.snap
    S, I, fill = D.S, D.I, D.fill
    years = [I(f"year_{y}") for y in YEARS]
    cal, capfix = I("year_2025"), I("cap_fixed_year")
    price, idle, mintr, r_, k2 = I("contract_price"), I("idle_share"), I("train_min_share"), I("ramp_years"), I("arith_k2")
    m2b, mw = I("m_to_b"), I("mw_to_gw")
    tk_gens = list(dict.fromkeys(snap["hdr_gen"]))
    base_cost = snap["hdr_cost"][1]
    lo_cost, hi_cost = snap["hdr_cost"][0], snap["hdr_cost"][2]
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    Dm, R = ctx.P["a2"]["D"], ctx.P["a2"]["R"]

    # ═════════════════════ Compute ═════════════════════
    C = Sheet(wb, "Compute", "C", "C", years,
              "Compute — 算力：加速器族 × 世代每 GW 年產能、逐合約供給 GW、自建、加速器組合、推論 GW（token 換算）、η、研發 GW、容量上限、VR 等值、敏感度",
              "規格 D6–D12（r1）：產能一律取 TK_Link（IF_TokGW_*、IF_Util、NonNV 產出比），不重算物理量；供給 GW 只由算力合約加總（D10 r1），機房租約只列對照；"
              "η＝2025 token 換算推論 GW ÷ 2025 推論支出換算 GW（D8）；研發 GW＝供給 ×（1 − 閒置）− 推論（D12，殘差）；容量上限 2025、2026 固定 1。",
              "GW＝IT 關鍵電力（口徑未明者依 Inputs 換算比例處理，基準視同 IT）。標「參數」的列只在 D 欄有值（與年度無關）。逐合約付款在 Cost 頁第二節。")
    C.section("一、世代參數（只在 D 欄；TK 世代名稱與成本情境為 SUMIFS 的文字鍵，取自快照表頭）")
    for k, g in enumerate(GENS):
        C.add(f"lab_{g}", f"參數：Tokenomics 世代名稱（{GEN_ZH[g]}）", "文字", [tk_gens[k]] + [None] * 5, "TK_HdrGen 文字鍵")
    C.add("lab_cost", "參數：Tokenomics 成本情境（基準）", "文字", [base_cost] + [None] * 5, "TK_HdrCost 文字鍵；每 GW 總產出三個成本情境同值")
    for g in GENS:
        for t, tz, _T in TIERS:
            C.add(f"tok_{g}_{t}", f"參數：每 GW 總產出 {GEN_ZH[g]}：{tz}（100% 利用率）", "M tok/GW/年",
                  only25(f"=SUMIFS(TK_IF_TokGW_{TK_TIER[t]},TK_HdrGen,«lab_{g}@$»,TK_HdrCost,«lab_cost@$»)"), f"TK IF_TokGW_{TK_TIER[t]}")

    C.section("二、加速器族參數（各年同值；NonNV 比例相對 VR200，D6；持有比 D9）")
    for f, fz, base, nnv, _N in FAMS:
        C.add(f"out_{f}", f"產出比（相對同代 NVIDIA）：{fz}", "倍", lambda i, c, nnv=nnv: f"=TK_NNV_{nnv}_Out" if nnv else f"=«vr_unit@$»",
              "TK NonNV 產出比（讀表）；NVIDIA＝VR200 等值係數的定義單位（＝1）")
        C.add(f"hold_{f}", f"持有比（相對 NVIDIA）：{fz}", "倍", lambda i, c, nnv=nnv: f"=TK_NNV_{nnv}_Hold" if nnv else f"=«vr_unit@$»", "TK NonNV 持有比（讀表）；NVIDIA＝1")
        C.add(f"price_{f}", f"每 GW 年合約價：{fz}", "$B/GW/年", lambda i, c, f=f: f"={price}*«hold_{f}»", "合約價（Inputs）× 持有比（D9）", name=f"CMP_Price_{FAM[f][4]}")
        C.add(f"vrf_{f}", f"VR200 等值係數（Sol 層級每 GW 產能比）：{fz}", "倍",
              lambda i, c, base=base, f=f: f"=«tok_{base}_mid@$»/«tok_vr200_mid@$»*«out_{f}»", "＝基準世代 Sol 產能 ÷ VR200 Sol 產能 × 產出比（同 OpenAI v0.6 VR 等值口徑）")
    # VR200 等值係數單位（＝VR200 Sol ÷ VR200 Sol）：作為 NVIDIA 族的產出比與持有比（避免公式內常數，E6）
    C.add("vr_unit", "參數：單位（VR200 Sol 產能 ÷ 自身）", "倍", only25("=«tok_vr200_mid@$»/«tok_vr200_mid@$»"), "恆等於 1；NVIDIA 族的產出比、持有比引用本格（E6）")

    C.section("三、既有雲端算力（2025 前簽：AWS Trainium2、Google TPU v6e、NVIDIA Hopper；2025＝說明書實際算力支出，D8／D14）")
    for f, key in LEGACY[:2]:
        C.add(f"lsh_{f}", f"既有雲端組合：{FAM[f][1]}", "比例", lambda i, c, key=key: f"={I(key)}", "Inputs（Assumed；D7）")
    C.add("lsh_tpu6", f"既有雲端組合：{FAM['tpu6'][1]}", "比例", lambda i, c: "=(1-«lsh_trn2»-«lsh_hopper»)", "＝1 − Trainium2 − Hopper")
    C.add("lprice", "既有雲端加權合約價", "$B/GW/年", lambda i, c: "=" + "+".join(f"«lsh_{f}»*«price_{f}»" for f, _ in LEGACY), "Σ 組合 × 各族合約價")
    C.add("sup_legacy", "供給 GW：既有雲端", "GW", lambda i, c: "=«K:pay_legacy»/«lprice»", "付款（Cost）÷ 加權合約價", name="CMP_Supply_Legacy")

    C.section("四、逐合約供給 GW（D10 r1：只計算力合約；金額型＝付款 ÷ 合約價；GW 型與雙揭露型＝揭露 GW × IT 換算 × 爬坡）")
    C.add("pue", "PUE（Tokenomics Inputs 頁；非具名讀表）", "倍", lambda i, c: "=TK_PUE", "TK_PUE（1.2；1.1–1.3）")
    C.add("itconv", "揭露 GW → IT GW 換算係數", "倍", lambda i, c: f"=(1-{I('unclear_facility_share')})+{I('unclear_facility_share')}/«pue»",
          "（1 − 視為設施口徑比例）＋ 比例 ÷ PUE；基準比例 0（視同 IT）")
    C.add("year", "年度（同第 6 列；爬坡公式用）", "年", lambda i, c: f"={c}$6", "")

    def act(c, s, e):
        return f"({c}$6>={s})*({c}$6<={e})"

    def ramp(c, s):
        return f"MIN({r_},{c}$6-{s}+1)/{r_}"

    for key, zh, fam, kind, p, _par, note in CONTRACTS:
        P = {k: fill(v) for k, v in p.items()}
        if kind in ("amt", "monthly"):
            fs = lambda i, c, key=key, fam=fam: f"=«K:pay_{key}»/«price_{fam}»"  # noqa: E731
        elif kind == "both":
            fs = lambda i, c, P=P: f"={P['G']}*«itconv»*{ramp(c, P['s'])}*{act(c, P['s'], P['e'])}"  # noqa: E731
        else:   # gw
            fs = lambda i, c, P=P: f"={P['G']}*«itconv»*{ramp(c, P['s'])}*({c}$6>={P['s']})*({c}$6<{P['s']}+{P['term']})"  # noqa: E731
        C.add(f"sup_{key}", f"供給 GW：{zh}", "GW", fs, note + f"；族＝{FAM[fam][1].split('（')[0]}", name=f"CMP_Supply_{key.capitalize()}")
    C.add("note_coreweave", "（未計入）CoreWeave（2026-04 多年合約；金額、容量皆未揭露）", "GW", [None] * 6,
          "找不到金額與容量（A1 not_found）→ 不計入（不杜撰）；列入資料缺口")

    C.section("五、自建（Fluidstack $50B；D11：自有資本支出自投產年起計入供給與命題分母）")
    fs_, fy = I("fluid_start"), I("fluid_years")
    C.add("own_capex", "自有資本支出（Fluidstack；均攤）", "$B", lambda i, c: f"={S('cmp_fluidstack_50b')}/{fy}*({c}$6>={fs_})*({c}$6<{fs_}+{fy})",
          "SRC_ANT_212 $50B ÷ 均攤年數（Inputs 3 年，2–5）；起始 2026", name="CMP_OwnedCapex")
    C.add("own_cum", "累計自有資本支出（至 年 − 投產落後年數）", "$B",
          lambda i, c: f"=SUMPRODUCT(($D$6:$I$6<={c}$6-{I('own_lag')})*(«own_capex@R»))", "落後 1 年（Inputs，同 OpenAI v0.6 INP_262）")
    C.add("own_cpg", "每 GW 資本支出（VR200 基準；TK IF_CapexTotal）", "$B/GW", lambda i, c: "=SUMIFS(TK_IF_CapexTotal,TK_HdrGen,«lab_vr200@$»,TK_HdrCost,«lab_cost@$»)",
          "同 OpenAI v0.6（IT 設備＋廠房）；第三方指 Fluidstack 站點用 Google TPU → 族＝TPU v7")
    C.add("sup_own", "供給 GW：自建（Fluidstack）", "GW", lambda i, c: "=«own_cum»/«own_cpg»", "累計資本支出 ÷ 每 GW 資本支出", name="CMP_Supply_Owned")

    C.section("六、供給 GW 合計與加速器組合（D7：由逐合約供給 GW 加權，時變）")
    ckeys = [c[0] for c in CONTRACTS]
    C.add("sup_con", "合約供給 GW（既有雲端＋新合約；不含自建）", "GW",
          lambda i, c: "=«sup_legacy»+" + "+".join(f"«sup_{k}»" for k in ckeys), "供應商持有成本的 GW", name="CMP_SupplyContractGW")
    C.add("sup", "供給 GW 合計（合約＋自建）", "GW", lambda i, c: "=«sup_con»+«sup_own»", "命題分母的實體 GW（不自動假設新增合約）", name="CMP_SupplyGW")
    for f, fz, *_r in FAMS:
        terms = [f"«sup_{k}»" for k in ckeys if CT[k][2] == f]
        if f in dict(LEGACY):
            terms.append(f"«sup_legacy»*«lsh_{f}»")
        if f == "tpu7":
            terms.append("«sup_own»")
        assert terms, f
        C.add(f"gw_{f}", f"GW：{fz.split('（')[0]}", "GW", lambda i, c, terms=terms: "=" + "+".join(terms),
              "Σ 該族合約 GW（既有雲端依組合拆分；自建歸 TPU v7）", name=f"CMP_GW_{FAM[f][4]}")
    C.add("gw_chk", "檢查：各族 GW 合計 − 供給 GW（應為 0）", "GW", lambda i, c: "=" + "+".join(f"«gw_{f}»" for f, *_ in FAMS) + "-«sup»", "")
    for f, fz, *_r in FAMS:
        C.add(f"s_{f}", f"組合占比：{fz.split('（')[0]}", "比例", lambda i, c, f=f: f"=«gw_{f}»/«sup»", "", name=f"CMP_Share_{FAM[f][4]}")
    C.add("s_sum", "組合占比合計（應＝1）", "比例", lambda i, c: "=" + "+".join(f"«s_{f}»" for f, *_ in FAMS), "Checks")

    C.section("七、層級組合（Demand 的 token ÷ 合計；付費、免費分開）")
    for kind, kz in (("paid", "付費"), ("free", "免費")):
        for t, tz, _T in TIERS:
            C.add(f"m_{kind}_{t}", f"{kz}：{tz}占比", "比例", lambda i, c, kind=kind, t=t: f"=«D:{kind}_{t}»/«D:{kind}_all»", "")

    C.section("八、每 GW 年產能（M tok/GW/年；基準利用率 TK IF_Util；式同 OpenAI v0.6：IF_Util ÷ Σ（層級占比 ÷ 該族該層級每 GW 總產出））")
    for kind, kz in (("paid", "付費"), ("free", "免費")):
        for g in GENS:
            C.add(f"cap_{kind}_{g}", f"{kz}：{GEN_ZH[g]}", "M tok/GW/年",
                  lambda i, c, kind=kind, g=g: "=TK_IF_Util/(" + "+".join(f"«m_{kind}_{t}»/«tok_{g}_{t}@$»" for t, _, _ in TIERS) + ")", "")
        for f, fz, base, nnv, _N in FAMS:
            if nnv:
                C.add(f"cap_{kind}_{f}", f"{kz}：{fz.split('（')[0]}", "M tok/GW/年", lambda i, c, kind=kind, f=f, base=base: f"=«out_{f}»*«cap_{kind}_{base}»",
                      "產出比 × 基準世代產能（D6；比例與層級無關）")
        C.add(f"cap_{kind}", f"{kz}：組合後每 GW 年產能", "M tok/GW/年",
              lambda i, c, kind=kind: "=" + "+".join(f"«s_{f}»*«cap_{kind}_{f}»" for f, *_ in FAMS), "Σ 組合占比 × 各族產能", name=f"CMP_Cap{kind.capitalize()}")

    C.section("九、推論 GW（token 換算）＝token ÷ 每 GW 年產能（付費、免費分開）")
    tm = I("t_to_m")
    C.add("tok_paid", "付費 token", "M tok/年", lambda i, c: f"=«D:paid_all»*{tm}", "Demand（T）× 10^6（Inputs 定義常數）；API 為計費 token")
    C.add("tok_free", "免費 token", "M tok/年", lambda i, c: f"=«D:free_all»*{tm}", "")
    C.add("gw_tpaid", "推論 GW（token 換算）：付費", "GW", lambda i, c: "=«tok_paid»/«cap_paid»", "", name="CMP_TokGW_Paid")
    C.add("gw_tfree", "推論 GW（token 換算）：免費", "GW", lambda i, c: "=«tok_free»/«cap_free»", "", name="CMP_TokGW_Free")
    C.add("gw_tok", "推論 GW（token 換算）：合計", "GW", lambda i, c: "=«gw_tpaid»+«gw_tfree»", "", name="CMP_TokGW")

    C.section("十、η 有效產出係數（D8：2025 推論支出校準）與有效推論 GW")
    C.add("e_spend", "2025 推論等非訓練算力支出（SRC，Derived＝說明書算力 7.33 − 訓練 4.1）", "$B", only25(f"={S('cmp_spend_2025_derived_inference')}"),
          "SRC_ANT_255（公式）；訓練 4.1 為 2025 年底預測（Interested-party），推論＝餘額（Derived）")
    C.add("e_spendgw", "2025 推論 GW（支出換算）＝推論支出 ÷ 既有雲端加權合約價", "GW", only25("=«e_spend»/«lprice»"), "", name="CMP_SpendGW2025", name_cols="D")
    C.add("eta25", "η（2025）＝token 換算 GW ÷ 支出換算 GW", "倍", only25("=«gw_tok»/«e_spendgw»"), "模型產能實際兌現的比例（Derived；D8）",
          name="CMP_Eta2025", name_cols="D")
    sw, tgt, ty = I("eta_switch"), I("eta_target"), I("eta_target_year")
    C.add("eta", "η（逐年）", "倍", lambda i, c: f"=«eta25@$»+{sw}*({tgt}-«eta25@$»)*MIN({ty}-{cal},{c}$6-{cal})/({ty}-{cal})",
          "開關 0（基準）＝沿用 2025；1＝線性回升至目標（D8 情境）", name="CMP_Eta")
    C.add("eff_paid", "有效推論 GW：付費", "GW", lambda i, c: "=«gw_tpaid»/«eta»", "")
    C.add("eff_free", "有效推論 GW：免費", "GW", lambda i, c: "=«gw_tfree»/«eta»", "")
    C.add("eff", "有效推論 GW（未截頂）", "GW", lambda i, c: "=«gw_tok»/«eta»", "2025＝支出換算 GW（校準恆等式）", name="CMP_InfGW_Eff")

    C.section("十一、研發 GW（D12：殘差）、閒置、容量上限（2025、2026 固定 1）")
    C.add("r_train25", "2025 訓練支出（SRC）", "$B", only25(f"={S('cmp_spend_2025_training_proj')}"), "SRC_ANT_253（2025 年底預測，Interested-party）")
    C.add("idle", "閒置 GW（校準年後）", "GW", lambda i, c: f"=«sup»*{idle}*({c}$6>{cal})", "Inputs 0.1（2026 起）")
    C.add("avail", "推論可用 GW＝供給 ×（1 − 閒置）×（1 − 訓練最低占比）", "GW", lambda i, c: f"=«sup»*(1-{idle})*(1-{mintr})", "", name="CMP_InfAvailGW")
    C.add("infcap", "推論 GW（截頂後）", "GW", lambda i, c: f"=IF({c}$6<={capfix},«eff»,MIN(«avail»,«eff»))", "2025、2026（≤ Inputs 固定年）＝有效推論；之後＝MIN（可用, 有效）",
          name="CMP_InfGW")
    C.add("rd", "研發 GW", "GW", lambda i, c: "=«r_train25»/«lprice»" if i == 0 else f"=«sup»*(1-{idle})-«infcap»",
          "2025＝訓練支出 ÷ 既有雲端加權合約價；2026 起＝供給 ×（1 − 閒置）− 推論（截頂後）", name="CMP_RDGW")
    C.add("dem", "總需求 GW＝有效推論（未截頂）＋研發", "GW", lambda i, c: "=«eff»+«rd»", "", name="CMP_DemandGW")
    C.add("cap", "容量上限係數（1＝未受限）", "倍", lambda i, c: "=«infcap»/«eff»", "接回 Revenue REV_CapFactor", name="CMP_CapFactor")
    C.add("flag", "缺口旗標（1＝有效推論超過可用 GW；固定年後）", "旗標", lambda i, c: f"=(«eff»>«avail»)*({c}$6>{capfix})", "不自動假設新增合約", name="CMP_CapFlag")
    C.add("ident", "檢查：推論（截頂後）＋研發＋閒置 − 供給", "GW", lambda i, c: "=«infcap»+«rd»+«idle»-«sup»", "各年應為 0（2025：說明書算力＝推論＋訓練）")
    C.add("rd_share", "研發占供給", "比例", lambda i, c: "=«rd»/«sup»", "")
    C.add("tk_rd", "對照：TK IF_AllocRDGW（研發 GW·年；Tokenomics 計畫當量，訓練世代 VR200）", "GW", lambda i, c: "=TK_IF_AllocRDGW", "D12 並列；口徑為單一前沿實驗室的物理下限")

    C.section("十二、VR 等值換算（命題分母；各族 Sol 產能比）")
    C.add("vrf", "機隊 VR 等值係數＝Σ 組合占比 × 各族 VR 等值係數", "倍", lambda i, c: "=" + "+".join(f"«s_{f}»*«vrf_{f}»" for f, *_ in FAMS), "", name="CMP_VReqFactor")
    C.add("vr_sup", "供給 VR 等值 GW", "GW", lambda i, c: "=«sup»*«vrf»", "", name="CMP_Supply_VReq")
    C.add("vr_inf", "推論 GW（截頂後）VR 等值", "GW", lambda i, c: "=«infcap»*«vrf»", "", name="CMP_InfGW_VReq")
    C.add("vr_rd", "研發 GW VR 等值", "GW", lambda i, c: "=«rd»*«vrf»", "", name="CMP_RDGW_VReq")
    C.add("vrf_alt", "對照：依付費層級組合加權的 VR 等值係數（付費組合後產能 ÷ VR200 付費產能）", "倍", lambda i, c: "=«cap_paid»/«cap_paid_vr200»", "替代口徑")

    C.section("十三、對照（只列差距，不校準）：報導年底 GW、揭露容量、機房租約（D10 r1：不另計 GW）")
    C.add("ye_model", "模型年底 GW 近似＝（本年 ＋ 次年年均供給）÷ 2", "GW", lambda i, c: None if i == 5 else f"=AVERAGE(«sup»,«sup@{YC[i + 1]}»)",
          "年均 → 年底的近似（線性）")
    C.add("ye_rep", "報導年底 GW（2025 約 1.4；2026 約 5；2027 約 10）", "GW",
          [f"={S('cmp_gw_ye_2025_fourweek')}", f"={S('cmp_gw_target_nyt')}", f"={S('cmp_gw_target_nyt_2027')}", None, None, None],
          "SRC_ANT_269、267、268（媒體轉述投資人資料；口徑未明）")
    C.add("ye_gap", "模型 ÷ 報導 − 1（|差距| > 門檻 → Checks WARN：2025、2026）", "比例", [f"=(«ye_model»-«ye_rep»)/«ye_rep»"] * 3 + [None] * 3, "",
          name="CMP_YEGap", name_cols="DF")
    C.add("ye_mai", "對照：Measured AI 年底估計（2025＝1.4；2026＝5.0）", "GW", [f"={S('cmp_gw_ye_2026_measuredai')}_Lo", f"={S('cmp_gw_ye_2026_measuredai')}"] + [None] * 4,
          "SRC_ANT_265（Analogy）")
    C.add("new148", "對照：2025-10 起新簽容量（The Information 彙整 ≥14.8 GW）vs 模型新合約 GW（不含既有雲端）", "GW",
          lambda i, c: f"=«sup»-«sup_legacy»" if i else f"={S('cmp_total_gw_148')}_Lo", "D 欄＝14.8（SRC_ANT_242）；E–I＝模型新合約＋自建")
    C.add("ref_google", "對照：Google TPU「2026 年上線遠超過 1 GW」vs 模型 Google TPU 2026 年均 GW", "GW", [f"={S('cmp_google_tpu_2025_gw')}_Lo", "=«sup_google»"] + [None] * 4,
          "D＝揭露下限；E＝模型（金額型）")
    C.add("ref_aws", "對照：Amazon「2026 年底前新增近 1 GW」／「最多 5 GW」vs 模型 Amazon GW", "GW",
          [f"={S('cmp_aws_2026_gw_ye2026')}", "=«sup_aws»", "=«sup_aws»", f"={S('cmp_aws_2026_gw_cap')}_Hi", None, None], "D＝近 1 GW；E、F＝模型；G＝上限 5 GW")
    C.add("ref_azure", "對照：Azure＋NVIDIA「最多 1 GW」vs 模型 Azure GW（2028 滿載）", "GW", [f"={S('cmp_msft_nvda_1gw')}_Hi", None, None, "=«sup_azure»", None, None], "")
    C.add("ref_spacex", "對照：SpaceX Colossus 1「>300 MW」vs 模型 SpaceX GW（月費 ÷ Hopper 合約價）", "GW",
          [f"={S('cmp_spacex_colossus1_mw')}_Lo/{mw}", "=«sup_spacex»", "=«sup_spacex»", None, None, None], "若以 0.3 GW 計，SpaceX 隱含每 GW 年付款約 $50B（見報告）")
    C.add("ref_rainier", "對照：Project Rainier IT 電力（Epoch 衛星估計 910 MW）vs 模型 2025 既有雲端 Trainium2 GW", "GW",
          [f"={S('cmp_aws_rainier_epoch_it_mw')}/{mw}", "=«sup_legacy@D»*«lsh_trn2@D»"] + [None] * 4, "D＝Epoch（2026-10 現況）；E＝模型 2025 年均")
    C.add("ref_rainier_fac", "對照：New Carlisle 園區用電 2.2 GW（設施）÷ PUE", "GW", [f"={S('cmp_aws_rainier_site_gw')}/«pue»"] + [None] * 5, "AWS 自有園區（設施口徑）")
    for key, zh, f, note in LEASES:
        C.add(f"lease_{key}", f"機房租約（不計入供給）：{zh}", "GW", [f"={fill(f)}/{mw}"] + [None] * 5, note)
    C.add("lease_sum", "機房租約容量合計（不計入供給；D10 r1）", "GW", ["=" + "+".join(f"«lease_{k}»" for k, *_ in LEASES)] + [None] * 5,
          "對照：2027–2028 新合約＋自建 GW（模型）是否足以容納這些機房")

    C.section("十四、敏感度（Excel 公式；引用 Inputs／TK 的值、低、高欄）：A 合約價、B η 路徑、C TPU／Trainium／AMD 產出比")
    rp = D.inp_row("contract_price")
    var = [k for k in ckeys if CT[k][3] in ("amt", "monthly")]
    fix = [k for k in ckeys if CT[k][3] in ("both", "gw")]
    C.add("sup_var", "供給 GW 中隨合約價反比者（金額型＋既有雲端）", "GW", lambda i, c: "=«sup_legacy»+" + "+".join(f"«sup_{k}»" for k in var), "")
    C.add("vr_var", "同上 VR 等值", "GW", lambda i, c: "=«sup_legacy»*(" + "+".join(f"«lsh_{f}»*«vrf_{f}»" for f, _ in LEGACY) + ")+"
          + "+".join(f"«sup_{k}»*«vrf_{CT[k][2]}»" for k in var), "")
    for tag, col, zh in (("lo", "F", "低端"), ("base", "E", "基準"), ("hi", "G", "高端")):
        P = f"Inputs!${col}${rp}"
        k = f"A_{tag}"
        C.add(f"{k}_p", f"A 合約價（{zh}）", "$B/GW/年", only25(f"={P}"), f"Inputs {price} 的{zh}欄")
        C.add(f"{k}_sup", f"A 供給 GW（合約價{zh}）", "GW", lambda i, c, k=k: f"=«sup»-«sup_var»+«sup_var»*{price}/«{k}_p@$»", "金額型與既有雲端 GW 與合約價成反比；GW 型、雙揭露型、自建不變")
        C.add(f"{k}_vr", f"A 供給 VR 等值 GW（合約價{zh}）", "GW", lambda i, c, k=k: f"=«vr_sup»-«vr_var»+«vr_var»*{price}/«{k}_p@$»", "")
        C.add(f"{k}_cc", f"A 算力成本（合約價{zh}）", "$B", lambda i, c, k=k: f"=«K:cc»+«K:pay_gwtype»*(«{k}_p@$»/{price}-«vr_unit@$»)",
              "GW 型付款與合約價成正比；金額型、雙揭露型、既有雲端、自建不變")
        C.add(f"{k}_eff", f"A 有效推論 GW（合約價{zh}）", "GW", lambda i, c, k=k: f"=«eff»*{price}/«{k}_p@$»", "η₂₀₂₅ 與合約價成正比（支出換算 GW 反比）")
        C.add(f"{k}_cap", f"A 容量上限係數（合約價{zh}）", "倍",
              lambda i, c, k=k: f"=IF({c}$6<={capfix},«vr_unit@$»,MIN(«{k}_sup»*(1-{idle})*(1-{mintr}),«{k}_eff»)/«{k}_eff»)", "")
        C.add(f"{k}_gap", f"A 每 VR 等值 GW 差額（含股權報酬；合約價{zh}）", "$B/GW/年",
              lambda i, c, k=k: f"=(«R:net»*«{k}_cap»-«{k}_cc»-«K:ncx»-«K:sbc»)/«{k}_vr»", "命題 1 的敏感度")
    for tag, zh, f in (("keep", "沿用（基準）", lambda i, c: "=«eta25@$»"),
                       ("rise", "線性回升至目標", lambda i, c: f"=«eta25@$»+({tgt}-«eta25@$»)*MIN({ty}-{cal},{c}$6-{cal})/({ty}-{cal})")):
        k = f"B_{tag}"
        C.add(f"{k}_eta", f"B η 路徑：{zh}", "倍", f, "")
        C.add(f"{k}_eff", f"B 有效推論 GW（η {zh}）", "GW", lambda i, c, k=k: f"=«gw_tok»/«{k}_eta»", "")
        C.add(f"{k}_cap", f"B 容量上限係數（η {zh}）", "倍",
              lambda i, c, k=k: f"=IF({c}$6<={capfix},«vr_unit@$»,MIN(«avail»,«{k}_eff»)/«{k}_eff»)", "供給不變")
    for tag, suf, zh in (("lo", "OutLo", "低端"), ("hi", "OutHi", "高端")):
        k = f"C_{tag}"
        nn = [(f, base, nnv) for f, _, base, nnv, _ in FAMS if nnv]
        for kind in ("paid", "free"):
            C.add(f"{k}_cap_{kind}", f"C 組合後每 GW 年產能（{'付費' if kind == 'paid' else '免費'}；TPU／Trainium／AMD 產出比皆取{zh}）", "M tok/GW/年",
                  lambda i, c, kind=kind: f"=«cap_{kind}»+" + "+".join(f"«s_{f}»*(TK_NNV_{nnv}_{suf}-«out_{f}»)*«cap_{kind}_{base}»" for f, base, nnv in nn), "")
        C.add(f"{k}_tok", f"C 推論 GW（token 換算；產出比{zh}）", "GW", lambda i, c, k=k: f"=«tok_paid»/«{k}_cap_paid»+«tok_free»/«{k}_cap_free»", "")
        C.add(f"{k}_eff", f"C 有效推論 GW（產出比{zh}）", "GW", lambda i, c, k=k: f"=«{k}_tok»/(«{k}_tok@$»/«e_spendgw@$»)", "η 以該情境 2025 token GW 重新校準（沿用路徑）")
        C.add(f"{k}_cap", f"C 容量上限係數（產出比{zh}）", "倍", lambda i, c, k=k: f"=IF({c}$6<={capfix},«vr_unit@$»,MIN(«avail»,«{k}_eff»)/«{k}_eff»)", "")
        C.add(f"{k}_vrf", f"C 機隊 VR 等值係數（產出比{zh}）", "倍",
              lambda i, c: "=«vrf»+" + "+".join(f"«s_{f}»*(TK_NNV_{nnv}_{suf}-«out_{f}»)*«vrf_{base}»" for f, base, nnv in nn), "")
        C.add(f"{k}_gap", f"C 每 VR 等值 GW 差額（含股權報酬；產出比{zh}）", "$B/GW/年",
              lambda i, c, k=k: f"=(«R:net»*«{k}_cap»-«K:full»)/(«sup»*«{k}_vrf»)", "")

    # ═════════════════════ Cost ═════════════════════
    K = Sheet(wb, "Cost", "K", "K", years,
              "Cost — 成本：逐合約實付＋自建資本支出（現金口徑）、供應商持有成本與雲端毛利、經濟口徑、非算力成本與股權報酬、推論／研發拆分、命題表（每 VR 等值 GW）、D22 對照",
              "規格 D10、D11、D13、D21、D22：算力成本＝合約實付（既有雲端 2025＝說明書實際 7.33）＋自建資本支出；TK IF_HoldEcon × 持有比只作供應商成本參考，差額＝雲端毛利；"
              "非算力：2025＝說明書營業費用 12.65 − 算力 7.33 − 平台抽成 0.351（抽成已自營收扣除，D5 r1），之後人數 × 每人成本；股權報酬單列（命題全成本含之）。",
              "命題表用營收淨額（截頂後）；2025 淨損 $42B 中約 $34B 非現金會計費用不在營業費用內（D21），不進命題表。$B＝十億美元；GW＝IT 實體 GW。")
    K.section("一、加速器族每 GW 年持有成本（經濟口徑；TK IF_HoldEcon × 持有比；只在 D 欄）")
    K.add("lab_lo", "參數：成本情境鍵（低成本）", "文字", [lo_cost] + [None] * 5, "")
    K.add("lab_hi", "參數：成本情境鍵（高成本）", "文字", [hi_cost] + [None] * 5, "")
    for f, fz, base, _n, _N in FAMS:
        for tag, lab in (("base", "«C:lab_cost@$»"), ("lo", "«lab_lo@$»"), ("hi", "«lab_hi@$»")):
            K.add(f"hold_{f}_{tag}", f"參數：每 GW 年持有成本（{ {'base': '基準', 'lo': '低成本', 'hi': '高成本'}[tag] }）：{fz.split('（')[0]}", "$B/GW/年",
                  only25(f"=SUMIFS(TK_IF_HoldEcon,TK_HdrGen,«C:lab_{base}@$»,TK_HdrCost,{lab})*«C:hold_{f}@D»"), "TK IF_HoldEcon（基準世代）× 持有比")

    K.section("二、算力成本（現金口徑）＝逐合約實付＋自建資本支出（D10、D11）")
    ramp_s = lambda s, e: f"((({e}-{s}+1)+({e}-{s}+1)-{r_}+1)/{k2})"  # noqa: E731  Σ 權重＝（2n − r ＋ 1）÷ 2
    others = [k for k in ckeys]
    K.add("act25", "2025 實際算力與基礎設施支出（說明書，SRC）", "$B", only25(f"={S('cmp_spend_2025_actual_prospectus')}"), "SRC_ANT_252（Interested-party）")
    K.add("pay_legacy", "合約實付：既有雲端（2025 前簽：AWS、Google、NVIDIA）", "$B",
          lambda i, c: ("=«act25»-(" + "+".join(f"«pay_{k}»" for k in others) + "+«C:own_capex»)") if i == 0
          else f"=«pay_legacy@D»*{I('legacy_carry_2026')}*({c}$6={I('year_2026')})",
          "2025＝實際算力支出 − 其他合約與自建（皆 0：新合約 2026 起）＝校準；2026＝2025 × 延續比例（Inputs 0.33）；2027 起 0（併入招股書新承諾）", name="COST_Pay_Legacy")
    for key, zh, fam, kind, p, _par, note in CONTRACTS:
        P = {k: fill(v) for k, v in p.items()}
        if kind in ("amt", "both"):
            fs = lambda i, c, P=P: f"={P['T']}/{ramp_s(P['s'], P['e'])}*MIN({r_},{c}$6-{P['s']}+1)/{r_}*({c}$6>={P['s']})*({c}$6<={P['e']})"  # noqa: E731
            nt = "總額 ÷ Σ 爬坡權重 × 當年權重（期間內；權重＝MIN(r, k)÷r）"
        elif kind == "monthly":
            fs = lambda i, c, P=P: (f"={P['F']}*(({c}$6={P['s']})*{P['m1']}+({c}$6={P['e']})*{P['m2']}"  # noqa: E731
                                    f"+({c}$6>{P['s']})*({c}$6<{P['e']})*{I('months_per_year')})")
            nt = "月費 × 當年計費月數"
        else:
            fs = lambda i, c, key=key, fam=fam: f"=«C:sup_{key}»*«C:price_{fam}»"  # noqa: E731
            nt = "GW × 合約價 × 持有比"
        K.add(f"pay_{key}", f"合約實付：{zh.split('（')[0]}", "$B", fs, nt, name=f"COST_Pay_{key.capitalize()}")
    K.add("pay", "合約實付合計", "$B", lambda i, c: "=«pay_legacy»+" + "+".join(f"«pay_{k}»" for k in ckeys), "", name="COST_ContractPay")
    K.add("pay_gwtype", "其中：GW 型合約付款（與合約價成正比；敏感度 A）", "$B", lambda i, c: "=" + "+".join(f"«pay_{k}»" for k in ckeys if CT[k][3] == "gw"), "AMD、Lambda")
    K.add("own", "自建資本支出（Fluidstack）", "$B", lambda i, c: "=«C:own_capex»", "Compute 第五節", name="COST_OwnedCapex")
    K.add("cc", "算力成本（現金口徑）＝合約實付＋自建資本支出", "$B", lambda i, c: "=«pay»+«own»", "命題輸出基準（D10）", name="COST_Compute")
    K.add("gap25", "差距：2025 算力成本 − 說明書實際（應為 0）", "$B", only25("=«cc»-«act25»"), "2025 校準（既有雲端列為餘額）", name="COST_Gap2025", name_cols="D")
    K.add("pay_gw", "每 GW 合約實付＝合約實付 ÷ 合約供給 GW", "$B/GW/年", lambda i, c: "=«pay»/«C:sup_con»", "")
    for key, zh, *_r in CONTRACTS:
        K.add(f"ppg_{key}", f"隱含每 GW 年付款：{zh.split('（')[0]}", "$B/GW/年", lambda i, c, key=key: f"=IF(«C:sup_{key}»>«C:year»-«C:year»,«pay_{key}»/«C:sup_{key}»,«C:year»-«C:year»)",
              "金額型、GW 型＝該族合約價；雙揭露型為揭露值的隱含價格")
    K.add("cmp_ratio", "算力成本 ÷ 營收總額", "倍", lambda i, c: "=«cc»/«R:gross»", "對照：2025 說明書約 1.6；2026Q1 0.71、Q2 預測 0.56（SRC_ANT_051、052）")
    K.add("h1_26", "對照：2026 上半年算力支出（報導）＝Q1 營收 × 0.71 ＋ Q2 營收 × 0.56", "$B",
          [None, f"={S('rev_2026_q1')}*{S('cost_compute_per_rev_2026_q1')}+{S('rev_2026_q2')}*{S('cost_compute_per_rev_2026_q2')}"] + [None] * 4,
          "費用口徑（Interested-party）；模型 2026 為現金口徑（含自建資本支出）", name="COST_H1Compute2026", name_cols="E")
    K.add("h1_26m", "對照：模型 2026 合約實付 ÷ 2 − 報導上半年", "$B", [None, f"=«pay»*{I('h2_months')}/{I('months_per_year')}-«h1_26»"] + [None] * 4,
          "只列差距（年內爬坡，上半年應低於全年一半）")

    K.section("三、供應商持有成本與雲端毛利（D10：持有成本＝合約 GW × TK IF_HoldEcon × 持有比；只作參考）")
    for tag, zh in (("base", "基準"), ("lo", "低成本"), ("hi", "高成本")):
        K.add(f"sh_{tag}", f"供應商持有成本（{zh}）", "$B",
              lambda i, c, tag=tag: "=" + "+".join(f"«C:gw_{f}»*«hold_{f}_{tag}@$»" for f, *_ in FAMS) + f"-«C:sup_own»*«hold_tpu7_{tag}@$»",
              "Σ 各族 GW × 每 GW 持有成本 − 自建（自有，不付供應商）", name="COST_SupplierHold" if tag == "base" else None)
    K.add("gm", "雲端毛利＝合約實付 − 供應商持有成本", "$B", lambda i, c: "=«pay»-«sh_base»", "", name="COST_CloudGM")
    K.add("gmp", "雲端毛利率", "比例", lambda i, c: "=«gm»/«pay»", "", name="COST_CloudGMPct")
    for tag, zh in (("lo", "低成本"), ("hi", "高成本")):
        K.add(f"gmp_{tag}", f"雲端毛利率（持有成本{zh}）", "比例", lambda i, c, tag=tag: f"=(«pay»-«sh_{tag}»)/«pay»", "TK 成本角落情境")

    K.section("四、經濟口徑（自建資本支出改以 TK 持有成本 × 自建 GW 年化）")
    K.add("own_hold", "自建持有成本（經濟）", "$B", lambda i, c: "=«C:sup_own»*«hold_tpu7_base@$»", "", name="COST_OwnedHold")
    K.add("ce", "算力成本（經濟口徑）＝合約實付＋自建持有成本", "$B", lambda i, c: "=«pay»+«own_hold»", "", name="COST_ComputeEcon")
    K.add("cash_econ", "現金口徑 − 經濟口徑", "$B", lambda i, c: "=«cc»-«ce»", "資本支出與持有成本的時間差")

    K.section("五、非算力成本（D13）與股權報酬：2025＝說明書；2026 起＝人數 × 每人年成本")
    K.add("opex25", "2025 營業費用總額（說明書；含算力）", "$B", only25(f"={S('cost_opex_2025')}"), "SRC_ANT_274")
    K.add("nc25", "2025 非算力營業費用（SRC，Derived＝營業費用 − 算力）", "$B", only25(f"={S('cost_noncompute_opex_2025')}"), "SRC_ANT_277（公式）")
    K.add("fee25", "其中：雲端平台抽成（說明書列行銷費用；本模型自營收扣除）", "$B", only25(f"={S('rev_partner_fee_2025')}"), "SRC_ANT_031；D5 r1、A2 預設 18")
    K.add("nc25x", "2025 非算力營業費用（扣平台抽成；含股權報酬）", "$B", only25("=«nc25»-«fee25»"), "避免與 REV_Net 重複扣除")
    K.add("hc", "員工人數（年均）", "人", lambda i, c: f"=AVERAGE({S('hc_revelio_2024q4')},{S('hc_revelio_2025q4')})" if i == 0 else f"=«hc@p»*(1+{I(f'hc_g_{YEARS[i]}')})",
          "2025＝2024Q4 與 2025Q4 平均（Revelio，Analogy）；2026 起 ×（1＋年增率，Inputs）", name="COST_Headcount")
    K.add("sbc_pp", "每人股權報酬", "$M/人/年", lambda i, c: f"={I('sbc_pp_2025')}" if i == 0 else f"=«sbc_pp@p»*(1+{I('sbc_pp_g')})", "Inputs（Assumed；未揭露）")
    K.add("sbc", "股權報酬（單列）", "$B", lambda i, c: f"=«hc»*«sbc_pp»/{m2b}", "人數 × 每人股權報酬", name="COST_SBC")
    K.add("ncx", "非算力營運費用（不含股權報酬）", "$B", lambda i, c: "=«nc25x»-«sbc»" if i == 0 else f"=«hc»*«pp»/{m2b}",
          "2025＝說明書（扣抽成）− 股權報酬；2026 起＝人數 × 每人年成本", name="COST_NonCompExSBC")
    K.add("pp", "每人年成本（不含股權報酬）", "$M/人/年", lambda i, c: f"=«ncx@D»*{m2b}/«hc@D»" if i == 0 else f"=«pp@p»*(1+{I('pp_cost_g')})",
          "2025＝Derived；2026 起 Inputs 年變動率")
    K.add("nci", "非算力成本（含股權報酬）", "$B", lambda i, c: "=«ncx»+«sbc»", "", name="COST_NonCompInclSBC")
    K.add("nc_rev", "非算力成本（含股權報酬）÷ 營收總額", "比例", lambda i, c: "=«nci»/«R:gross»", "")
    K.add("hc_ref", "對照：報導人數（2026Q1 Revelio 3,950；2026-06 公司口述約 3,500）", "人", [None, f"={S('hc_revelio_2026q1')}"] + [None] * 4, "SRC_ANT_286、295")
    K.add("emp_ref", "對照：2026Q1 員工費用 × 4（SRC Derived，約 $1.25B／季）vs 模型 2026 非算力（含股權報酬）", "$B",
          [None, f"={S('cost_employee_q1_2026_derived')}*{I('quarters_per_year')}"] + [None] * 4, "員工費用只是非算力成本的一部分")
    K.add("oploss25", "對照：2025 營業結果（模型）＝營收總額 − 算力 − 非算力（含抽成與股權報酬）", "$B", only25("=«R:gross»-«cc»-«nc25»"),
          "說明書營業虧損 8.06（SRC_ANT_278）；差額為四捨五入")
    K.add("oploss25g", "差距：模型 2025 營業結果 ＋ 說明書營業虧損", "$B", only25(f"=«oploss25»+{S('cost_oploss_2025')}"), "應接近 0", name="COST_OpLossGap2025",
          name_cols="D")

    K.section("六、推論／研發算力成本拆分（2025＝說明書 D8；2026 起依 GW 比例）")
    K.add("c_inf", "推論算力", "$B", lambda i, c: "=«C:e_spend»" if i == 0 else "=«cc»*«C:infcap»/«C:sup»", "2025＝SRC 推導 3.23；2026 起＝算力 × 推論 GW ÷ 供給", name="COST_InfCompute")
    K.add("c_rd", "研發算力", "$B", lambda i, c: "=«C:r_train25»" if i == 0 else "=«cc»*«C:rd»/«C:sup»", "2025＝SRC 4.1；2026 起＝算力 × 研發 GW ÷ 供給", name="COST_RDCompute")
    K.add("c_idle", "閒置算力", "$B", lambda i, c: "=«cc»*«C:idle»/«C:sup»", "")
    K.add("c_ck", "檢查：推論＋研發＋閒置 − 算力成本", "$B", lambda i, c: "=«c_inf»+«c_rd»+«c_idle»-«cc»", "應為 0")
    K.add("c_ref", "對照：公司 2026 預測（訓練 12、推論 7；2026-01，Interested-party）", "$B",
          [None, f"={S('cmp_spend_2026_training_proj')}+{S('cmp_spend_2026_inference_proj')}"] + [None] * 4, "SRC_ANT_258、259；只對照（推論 7 已過時，見 A1 筆記）")

    K.section("七、命題表（每 VR 等值 GW；分母＝供給 VR 等值 GW，含自建）：營收淨額、算力、非算力、全成本（含／不含股權報酬）、差額、覆蓋率")
    K.add("p_den", "供給 VR 等值 GW（分母）", "GW", lambda i, c: "=«C:vr_sup»", "CMP_Supply_VReq")
    K.add("p_rev", "營收淨額（截頂後）", "$B", lambda i, c: "=«R:net_c»", "REV_NetCapped（扣雲端平台抽成）")
    K.add("full", "全成本（含股權報酬；現金口徑）", "$B", lambda i, c: "=«cc»+«ncx»+«sbc»", "命題基準（較保守）", name="COST_FullCash")
    K.add("fullx", "全成本（不含股權報酬；現金口徑）", "$B", lambda i, c: "=«cc»+«ncx»", "", name="COST_FullCashExSBC")
    K.add("gap_b", "差額＝營收淨額 − 全成本（含股權報酬）", "$B", lambda i, c: "=«p_rev»-«full»", "負＝缺口（A4 融資）", name="COST_GapCash")
    K.add("p_rev_vr", "每 VR 等值 GW：營收淨額", "$B/GW/年", lambda i, c: "=«p_rev»/«p_den»", "", name="COST_PropRev_VR")
    K.add("p_cc_vr", "每 VR 等值 GW：算力成本（現金）", "$B/GW/年", lambda i, c: "=«cc»/«p_den»", "", name="COST_PropCompute_VR")
    K.add("p_nc_vr", "每 VR 等值 GW：非算力成本（不含股權報酬）", "$B/GW/年", lambda i, c: "=«ncx»/«p_den»", "", name="COST_PropNonComp_VR")
    K.add("p_sbc_vr", "每 VR 等值 GW：股權報酬", "$B/GW/年", lambda i, c: "=«sbc»/«p_den»", "", name="COST_PropSBC_VR")
    K.add("p_full_vr", "每 VR 等值 GW：全成本（含股權報酬）", "$B/GW/年", lambda i, c: "=«p_cc_vr»+«p_nc_vr»+«p_sbc_vr»", "", name="COST_PropFull_VR")
    K.add("p_fullx_vr", "每 VR 等值 GW：全成本（不含股權報酬）", "$B/GW/年", lambda i, c: "=«p_cc_vr»+«p_nc_vr»", "", name="COST_PropFullExSBC_VR")
    K.add("p_gap_vr", "每 VR 等值 GW：差額（含股權報酬）", "$B/GW/年", lambda i, c: "=«p_rev_vr»-«p_full_vr»", "命題 1：負＝營收不足以覆蓋全成本", name="COST_PropGap_VR")
    K.add("p_gapx_vr", "每 VR 等值 GW：差額（不含股權報酬）", "$B/GW/年", lambda i, c: "=«p_rev_vr»-«p_fullx_vr»", "", name="COST_PropGapExSBC_VR")
    K.add("p_cov", "覆蓋率＝營收淨額 ÷ 全成本（含股權報酬）", "倍", lambda i, c: "=«p_rev»/«full»", "", name="COST_Coverage")
    K.add("p_covx", "覆蓋率（不含股權報酬）", "倍", lambda i, c: "=«p_rev»/«fullx»", "", name="COST_CoverageExSBC")
    K.add("p_rev_gw", "對照：每實體 GW 營收淨額", "$B/GW/年", lambda i, c: "=«p_rev»/«C:sup»", "分母改用實體 GW")
    K.add("p_gap_gw", "對照：每實體 GW 差額（含股權報酬）", "$B/GW/年", lambda i, c: "=«gap_b»/«C:sup»", "")
    K.add("e_cc_vr", "經濟口徑：每 VR 等值 GW 算力成本", "$B/GW/年", lambda i, c: "=«ce»/«p_den»", "")
    K.add("e_gap_vr", "經濟口徑：每 VR 等值 GW 差額（含股權報酬）", "$B/GW/年", lambda i, c: "=«p_rev_vr»-«e_cc_vr»-«p_nc_vr»-«p_sbc_vr»", "", name="COST_PropGapEcon_VR")
    K.add("e_gap_b", "經濟口徑：差額（含股權報酬）", "$B", lambda i, c: "=«p_rev»-«ce»-«ncx»-«sbc»", "", name="COST_GapEcon")

    K.section("八、TK 單位成本參考（IF_FullCostDefault_Sol 為基準；只作參考，算力成本仍＝合約實付）")
    d2b = I("usd_to_b")
    for nm_, zh, key in (("IF_FullCostDefault_Sol", "下游預設 IF_FullCostDefault_Sol；基準", "tk_def"), ("IF_FullCost_Sol", "IF_FullCost_Sol；並列", "tk_full")):
        K.add(key, f"TK 參考：每 VR200 GW 年全成本（{zh}）", "$B/GW/年",
              lambda i, c, nm_=nm_: f"=SUMIFS(TK_{nm_},TK_HdrGen,«C:lab_vr200@$»,TK_HdrCost,«C:lab_cost@$»)*«C:tok_vr200_mid@$»*TK_IF_Util/{d2b}",
              "TK $/M（VR200 Sol 基準欄）× 每 GW 年產出（Sol）× 基準利用率 ÷ 10⁹")
    K.add("tk_gap", "差距：每 VR 等值 GW 算力成本（現金）− TK 參考（下游預設）", "$B/GW/年", lambda i, c: "=«p_cc_vr»-«tk_def»", "")

    K.section("九、D22 對照：2026 Q2 首次調整後營業利益（說明書／報導；不校準）")
    K.add("adj_oi", "模型 2026 調整後營業結果（近似）＝營收淨額 − 合約實付 − 非算力（不含股權報酬）", "$B", [None, "=«R:net»-«pay»-«ncx»"] + [None] * 4,
          "調整後口徑排除股權報酬；算力用合約實付（費用口徑近似，不含自建資本支出）", name="COST_AdjOI2026", name_cols="E")
    K.add("adj_oim", "模型 2026 調整後營業利益率", "比例", [None, "=«adj_oi»/«R:net»"] + [None] * 4, "", name="COST_AdjOIM2026", name_cols="E")
    K.add("adj_rep", "報導：2026 Q2 營業利益預測 $0.559B ÷ Q2 營收預測 $10.9B（實際營收 >$11.5B 且調整後營業利益為正）", "比例",
          [None, f"={S('proj_opinc_2026q2_wsj')}/{S('proj_rev_2026q2_wsj')}"] + [None] * 4, "SRC_ANT_348、347、350（Interested-party）", name="COST_AdjOIMRep2026Q2", name_cols="E")

    K.section("十、檢查列（各年應為 0；Checks 取最大絕對值）")
    K.add("ck_pay", "逐合約實付加總 − 合約實付合計", "$B", lambda i, c: "=«pay_legacy»+" + "+".join(f"«pay_{k}»" for k in ckeys) + "-«pay»", "")
    K.add("ck_amt", "金額型合約：|GW × 合約價 − 實付| 合計", "$B",
          lambda i, c: "=" + "+".join(f"ABS(«C:sup_{k}»*«C:price_{CT[k][2]}»-«pay_{k}»)" for k in var) + "+ABS(«C:sup_legacy»*«C:lprice»-«pay_legacy»)", "D10：GW＝金額 ÷ 合約價")
    K.add("ck_cc", "算力成本（現金）−（合約實付＋自建資本支出）", "$B", lambda i, c: "=«cc»-«pay»-«own»", "")
    K.add("ck_nc", "全成本 −（算力＋非算力（不含股權報酬）＋股權報酬）", "$B", lambda i, c: "=«full»-«cc»-«ncx»-«sbc»", "")

    K.section("十一、合約總額對帳（約十年；chat 端預設：逐合約加總對招股書約 $518B，差距 >25% → WARN）")
    for key, zh, fam, kind, p, par, _note in CONTRACTS:
        P = {k: fill(v) for k, v in p.items()}
        if kind in ("amt", "both"):
            f = f"={P['T']}"
            nt = "揭露總額"
        elif kind == "monthly":
            f = f"={P['F']}*({P['m1']}+{P['m2']}+({P['e']}-({P['s']}+1))*{I('months_per_year')})"
            nt = "月費 × 總月數（S-1 期間）"
        else:
            f = f"=«C:gw_unit_{key}@$»*«C:price_{fam}@D»*(({P['term']})+({P['term']})-{r_}+1)/{k2}"
            nt = "滿載年付款 × Σ 爬坡權重（期間 Inputs）"
        K.add(f"tot_{key}", f"合約總額（模型）：{zh.split('（')[0]}" + (f"［招股書：{par}］" if par else ""), "$B", only25(f), nt)
    K.add("tot6", "模型：招股書六家合作方對應合約加總", "$B", only25("=" + "+".join(f"«tot_{k}»" for k, *_r, par, _n in CONTRACTS if par)), "", name="COST_Commit6", name_cols="D")
    K.add("tot_all", "模型：全部新合約加總（含 Nscale、Lambda、Volta、Akamai）", "$B", only25("=" + "+".join(f"«tot_{k}»" for k in ckeys)), "", name="COST_CommitAll",
          name_cols="D")
    for nmz, key, suf in PARTNERS:
        K.add(f"pros_{key}", f"招股書：{nmz}", "$B", only25(f"={S(key)}{suf}"), "SRC（Interested-party）")
    K.add("pros_sum", "招股書六家合計（逐項）", "$B", only25("=" + "+".join(f"«pros_{k}»" for _n, k, _s in PARTNERS)), "")
    K.add("pros_total", "招股書總承諾（約，SRC_ANT_245）", "$B", only25(f"={S('cmp_prospectus_total_518b')}"), "")
    K.add("commit_gap", "差距：模型六家加總 ÷ 招股書總承諾 − 1（|差距| > 門檻 → WARN）", "比例", only25("=(«tot6»-«pros_total»)/«pros_total»"), "",
          name="COST_CommitGap", name_cols="D")
    K.add("commit_gap_all", "對照：模型全部新合約 ÷ The Information 彙整 $517B − 1", "比例", only25(f"=(«tot_all»-{S('cmp_total_commit_517b')}_Hi)/{S('cmp_total_commit_517b')}_Hi"), "")

    # Compute 需要的「GW 型合約滿載 GW」參數列（合約總額對帳用）
    for key, zh, fam, kind, p, _par, _note in CONTRACTS:
        if kind == "gw":
            C.add(f"gw_unit_{key}", f"參數：GW 型合約滿載 GW（{zh.split('（')[0]}）", "GW", only25(f"={fill(p['G'])}*«itconv@D»"), "揭露 GW ×（計入比例）× IT 換算")

    # ── 解析與 Revenue 容量上限列
    C.resolve()
    K.resolve()
    rcap = R.rows["cap"]
    for i, c in enumerate(YC):
        put(R.ws, f"{c}{rcap}", f"=Compute!{c}{C.rows['cap']}", fmt="0.000")
    # 頁序：Compute、Cost 放在 Revenue 之後
    names = wb.sheetnames
    for sh in ("Cost", "Compute"):
        wb.move_sheet(sh, offset=names.index("Revenue") + 1 - wb.sheetnames.index(sh))
    for sh, wB in ((C.ws, 66), (K.ws, 66)):
        sh.column_dimensions["A"].width = 7
        sh.column_dimensions["B"].width = wB
        sh.column_dimensions["C"].width = 12
        for c in YC:
            sh.column_dimensions[c].width = 12
        sh.column_dimensions["K"].width = 80
        sh.freeze_panes = "D7"
    ctx.P["a3"] = dict(C=C, K=K)


def checks(ctx):
    D = ctx.D
    S, I = D.S, D.I
    C, K = ctx.P["a3"]["C"], ctx.P["a3"]["K"]
    cr, kr = C.rows, K.rows
    yr = lambda sh, key: f"{sh.name}!$D${sh.rows[key]}:$I${sh.rows[key]}"  # noqa: E731
    mx = lambda sh, key: f"=MAX(MAX({yr(sh, key)}),-MIN({yr(sh, key)}))"  # noqa: E731
    d = lambda sh, key, col="D": f"{sh.name}!${col}${sh.rows[key]}"  # noqa: E731
    lmix = f"({I('legacy_share_trn2')}<0)+({I('legacy_share_hopper')}<0)+((1-{I('legacy_share_trn2')}-{I('legacy_share_hopper')})<0)"
    terms = "+".join(f"(({fill_(D, p['e'])}-{fill_(D, p['s'])}+1)<{I('ramp_years')})" for k, _z, _f, kind, p, *_ in CONTRACTS if kind in ("amt", "both"))
    return [
        ("a3_cc25", "2025 算力成本：模型 − 說明書實際 $7.33B（$B）", "=COST_Gap2025", 0, "tol", "既有雲端列為 2025 餘額（校準）；其餘合約 2026 起"),
        ("a3_split25", "2025 推論 3.23 ＋ 訓練 4.1 − 實際 7.33（$B）", f"={S('cmp_spend_2025_derived_inference')}+{S('cmp_spend_2025_training_proj')}-{S('cmp_spend_2025_actual_prospectus')}",
         0, "tol", "D8：η 校準用的推論支出＝實際 − 訓練（Derived）"),
        ("a3_sup_ge_inf", "供給 GW < 推論 GW（截頂後）的年數", f"=SUMPRODUCT(--({yr(C, 'sup')}<{yr(C, 'infcap')}))", 0, "eq", "工作單 A3 第 3 步：供給 ≥ 推論"),
        ("a3_rd_neg", "研發 GW < 0 的年數（推論吃掉研發；WARN）", f'=COUNTIF({yr(C, "rd")},"<0")', "0", "warn", "D12 殘差：2026 容量上限固定 1 時可能發生"),
        ("a3_eta", "η 超出合理範圍（Inputs 0.5–2）的年數（WARN）",
         f'=COUNTIF({yr(C, "eta")},"<"&{I("eta_warn_lo")})+COUNTIF({yr(C, "eta")},">"&{I("eta_warn_hi")})', "0", "warn", "工作單 A3 第 3 步"),
        ("a3_ident", "推論（截頂後）＋研發＋閒置 − 供給（各年最大絕對差，GW）", mx(C, "ident"), 0, "tol", "2025：說明書算力＝推論＋訓練；2026 起：研發為殘差"),
        ("a3_gwchk", "各族 GW 合計 − 供給（各年最大絕對差）", mx(C, "gw_chk"), 0, "tol", ""),
        ("a3_mix", "組合占比合計 − 1（各年最大絕對差）", f"=MAX(MAX({yr(C, 's_sum')}),-MIN({yr(C, 's_sum')}))-{d(C, 'vr_unit')}", None, "info",
         "合計恆為 1（本列＝最大值 − 1，只列示）"),
        ("a3_lmix", "既有雲端組合占比 < 0 或餘額 < 0 的違反數", "=" + lmix, 0, "eq", "Inputs 只設 Trainium2、Hopper"),
        ("a3_term", "爬坡型合約期間短於爬坡年數的合約數", "=" + terms, 0, "eq", "Σ 權重公式（2n − r ＋ 1）÷ 2 需 n ≥ r"),
        ("a3_cap2526", "容量上限係數 2025、2026 ≠ 1 的年數", f"=({d(C, 'cap')}<>{d(C, 'vr_unit')})+({d(C, 'cap', 'E')}<>{d(C, 'vr_unit')})", 0, "eq",
         "chat 端交接：校準年不截頂"),
        ("a3_revcap", "Revenue 容量上限係數 ≠ Compute 的年數", "=SUMPRODUCT(--(REV_CapFactor<>CMP_CapFactor))", 0, "eq", ""),
        ("a3_pay", "逐合約實付加總 − 合計（各年最大絕對差）", mx(K, "ck_pay"), 0, "tol", ""),
        ("a3_amt", "金額型合約 GW × 合約價 − 實付（各年最大）", mx(K, "ck_amt"), 0, "tol", "D10"),
        ("a3_cc", "算力成本 −（合約實付＋自建）（各年最大）", mx(K, "ck_cc"), 0, "tol", ""),
        ("a3_full", "全成本 − 各項合計（各年最大）", mx(K, "ck_nc"), 0, "tol", ""),
        ("a3_csplit", "推論＋研發＋閒置算力 − 算力成本（各年最大）", mx(K, "c_ck"), 0, "tol", ""),
        ("a3_ncx25", "2025 非算力（不含股權報酬）≤ 0（股權報酬假設超過費用）", f"=IF({d(K, 'ncx')}<=0,1,0)", 0, "eq", "每人股權報酬 × 人數 > 說明書非算力費用時轉 ERR"),
        ("a3_commit", "逐合約加總（招股書六家對應）÷ 招股書 $518B − 1（|差距| > 門檻 → WARN）", "=COST_CommitGap", I("commit_warn_threshold"), "warn",
         "chat 端預設；機房租約與 Nscale、Lambda、Volta、Akamai 不在招股書六家之內"),
        ("a3_ye25", "2025 年底 GW：模型近似 ÷ 報導 1.4 − 1（WARN）", f"=INDEX(CMP_YEGap,1,1)", I("gw_warn_threshold"), "warn", "只對照、不校準"),
        ("a3_ye26", "2026 年底 GW：模型近似 ÷ 報導 5 − 1（WARN）", f"=INDEX(CMP_YEGap,1,2)", I("gw_warn_threshold"), "warn", "只對照、不校準"),
        ("a3_ye27", "對照：2027 年底 GW 模型近似 ÷ 報導約 10 − 1", f"=INDEX(CMP_YEGap,1,3)", None, "info", "NYT「再倍增」（Derived）"),
        ("a3_oploss25", "對照：2025 營業結果 模型 − 說明書（$B；四捨五入差）", "=COST_OpLossGap2025", None, "info", "D21：$34B 非現金費用不在營業費用內"),
        ("a3_d22m", "D22 對照：模型 2026 調整後營業利益率（近似）", "=COST_AdjOIM2026", None, "info", "不校準"),
        ("a3_d22r", "D22 對照：報導 2026 Q2 營業利益率（預測 0.559 ÷ 10.9；實際為正）", "=COST_AdjOIMRep2026Q2", None, "info", "SRC_ANT_347、348、350"),
        ("a3_h1", "對照：2026 上半年算力支出（報導比率推得，$B）", "=COST_H1Compute2026", None, "info", "費用口徑；模型 2026 合約實付見 Cost 第二節"),
        ("a3_eta25", "結果：η（2025）", "=CMP_Eta2025", None, "info", "D8"),
        ("a3_gap30", "結果：2030 每 VR 等值 GW 差額（含股權報酬，$B/GW/年）", "=INDEX(COST_PropGap_VR,1,6)", None, "info", "命題 1"),
        ("a3_cov30", "結果：2030 覆蓋率（含股權報酬）", "=INDEX(COST_Coverage,1,6)", None, "info", ""),
        ("a3_oaigap30", "OpenAI 並排：2030 每 VR 等值 GW 差額 Anthropic − OpenAI", "=INDEX(COST_PropGap_VR,1,6)-INDEX(OAI_COST_PropGap_VR,1,6)", None, "info", "OAI_Link（D20）"),
        ("a3_oaicov30", "OpenAI 並排：2030 覆蓋率 Anthropic ÷ OpenAI", "=INDEX(COST_Coverage,1,6)/INDEX(OAI_COST_Coverage,1,6)", None, "info", ""),
        ("a3_oaisup30", "OpenAI 並排：2030 供給 VR 等值 GW Anthropic ÷ OpenAI", "=INDEX(CMP_Supply_VReq,1,6)/INDEX(OAI_CMP_Supply_VReq,1,6)", None, "info", ""),
        ("a3_oaicc30", "OpenAI 並排：2030 算力成本 Anthropic ÷ OpenAI", "=INDEX(COST_Compute,1,6)/INDEX(OAI_COST_Compute,1,6)", None, "info", ""),
    ]


def fill_(D, tpl):
    return D.fill(tpl)

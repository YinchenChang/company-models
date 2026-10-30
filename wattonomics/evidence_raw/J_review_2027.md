# Independent review of candidates_v0.1.5.csv (31 rows): Wattonomics T2

Reviewer: independent agent that did not gather this evidence. Date: 2026-10-31.
Method: I opened every URL myself with WebFetch before reading the gatherer notes (b1/b2_candidates.csv, epoch_summary.md). For Epoch pages I copied the timeline and computed each increment myself. For dedupe I checked against T2_current_v0.1.4.csv.
Verdicts: 屬實 = accurate; 小誤 = minor error (corrected value given); 不符 = contradicted; 無法核對 = could not verify.

| n | key | verdict | corrected mw / basis / date / op | recommend inc | quote (<=25 words) | url opened |
|---|---|---|---|---|---|---|
| 1 | cp_googl_ftw1 | 屬實 | none (345−173=172, IT, 2027-03, googl) | 1 | Dec 1 2026 "Building 2 operational (estimated)" 173; Mar 31 2027 "Building 3 operational (estimated)" 345 (projected) | epoch google-fort-wayne |
| 2 | cp_googl_ftw2 | 屬實 | none (518−345=173, 2027-08) | 1 | Aug 28 2027 "Building 4 operational (estimated)" 518 MW (projected) | epoch google-fort-wayne |
| 3 | cp_googl_cdr1 | 屬實 | none (174−87=87, 2027-02) | 1 | Dec 31 2026 Building 1 87; Feb 22 2027 Building 2 174 (projected) | epoch google-cedar-rapids |
| 4 | cp_googl_cdr2 | 屬實 | none (261−174=87, 2027-05) | 1 | May 23 2027 Building 3 261 (projected); "Building 3 may trail Building 2 by about three months" | epoch google-cedar-rapids |
| 5 | cp_googl_sto1 | 屬實 | none (300−161=139, 2027-06) | 1 | Jun 30 2027 "Building 3 estimated operational" 300; projection "based on generator counts from an air quality permit application" | epoch google-storey-county |
| 6 | cp_googl_col5 | 屬實 | none (386−303=83, 2027-07) | 1 | 138 kV Hartman Farms line "expands the site by 100 MW … in-service date of July 1, 2027"; Jul 1 2027 386 | epoch google-columbus |
| 7 | cp_googl_lin2 | 屬實 | none (283−141=142, 2027-10) | 1 | May 12 2026 Building 1 141; Oct 1 2027 Building 2 283, "Estimated from Building 1's construction timeline" | epoch google-lincoln |
| 8 | cp_googl_gn1 | 屬實 | none (103 IT, 2027-03, googl). The claim mixes two things: Energy Mag's "Building 4 due March 2027" is a construction-finish date, while Epoch's 103 is Building 1 going operational. | 1 | Mar 15 2027 Building 1 103 MW; Crusoe "announced that it is the developer of Google's new data center campus" (w.media 2026-09-29) | epoch goodnight; theenergymag 2026-07-31; w.media |
| 9 | cp_googl_gn2 | 屬實 | none (206−103=103, 2027-05) | 1 | May 15 2027 "Building 2 operational, about two months after Building 1" 206 | epoch goodnight |
| 10 | cp_googl_gn3 | 小誤 | mw 800 → **412** (Buildings 3–6 at Building 1–2 size, about 103 each). IT, 2027-10, googl. Date carries downside risk: Building 6 construction finishes 2027-07-30, and Epoch's own lag for Building 1 was 6 months, which would put it near 2028-01. | 1 (corrected value only) | Abatement: "four 25-megawatt data halls planned for each building"; first phase "as much as 530 megawatts"; Epoch gives no per-building MW | epoch goodnight; theenergymag |
| 11 | cp_amzn_nalb | 屬實 (with a basis flag) | none (338−213=125, 2027-07). This figure covers only the AI chips, about 44% of IT. Total IT would be about 284 (see G2). | 1 | "We only include the capacity taken up by leading AI chips, which we estimate to be 44%." | epoch aws-new-albany |
| 12 | cp_amzn_rdg_b5 | 屬實 | none (284−228=56, 2027-02). No overlap with cp_amzn_rdg_b614 (840−284). | 1 | May 15 2026 Buildings 1-4 228; Feb 9 2027 "Building 5 operational (estimate)" 284; Sep 29 2028 840 | epoch amazon-ridgeland |
| 13 | cp_amzn_brw | 屬實 | none (169−108=61, 2027-08) | 1 | Aug 1 2027 Building 4 169, "modeled at twice Building 2's capacity" | epoch aws-berwick |
| 14 | cp_amzn_cip4 | 屬實 | none (81 IT, 2027-02, amzn, Interested-party). This is a subset of cpx_amzn_cip (216, inc 0), so there is no double count while that row stays at inc 0. | 1 | "81 IT MW on Feb. 28th, 2027 Data Hall 4"; "216 MW Critical IT Load … leased to Amazon Data Services, Inc." | sec.gov ex99.1 (Feb 2026) |
| 15 | cp_meta_mont1 | 屬實 | none (204−153=51, 2027-02) | 1 | Feb 2 2027 Buildings 5 and 6 204 (projected) | epoch meta-montgomery |
| 16 | cp_meta_mont2 | 屬實 | none (358−204=154, 2027-02) | 1 | Feb 23 2027 Buildings 3 and 4 358, "8-month lag behind Buildings 1 and 2" | epoch meta-montgomery |
| 17 | cp_meta_chy2 | 屬實 | none (207−152=55, 2027-02) | 1 | Jul 25 2026 152; Feb 15 2027 Buildings 1–4 207 (projected) | epoch meta-cheyenne |
| 18 | cp_meta_bg1 | 屬實 | none (86, 2027-02; no earlier MW point) | 1 | Feb 1 2027 Building 1 86 (projected) | epoch meta-bowling-green |
| 19 | cp_meta_bg2 | 屬實 | none (163−86=77, 2027-04) | 1 | Apr 1 2027 Building 2 163 (projected) | epoch meta-bowling-green |
| 20 | cp_orcl_sal_b1 | 屬實 (Derived) | none (461 設施, 2027-12, orcl/oai). Exhibit A-1 shows 1,383 as the monthly committed MW. Full ramp is 2027-12-01. Oracle confirms three buildings. The 1/3 split is an assumption, and a conservative one against the contract. | 1 | "scheduled to begin materializing in January 2027 with max load being achieved by December 2027"; delay "by up to 12 months in the aggregate" | npr mpsc-dte.pdf; oracle.com PR 2026-06-01; epoch openai-stargate-michigan |
| 21 | cp_orcl_redoak | 小誤 | Split the row: **60 IT @2027-01/02** (DFW10; TDLR TFO completion 2027-01-31) + **60 IT @2028-01** (DFW11; TDLR completion 2028-01-15). Only 60 falls in 2027. op orcl is moderate confidence (second-hand). | 1 (60 @2027) | DFW10 TFO "Estimated Completion Date: 1/31/2027"; DFW11 "1/15/2028"; "Oracle will be the tenant for all three" (per Business Insider) | dallasinnovates ?p=286676, ?p=290330; tdlr TABS2026015329, TABS2025016916; DCD |
| 22 | cp_nbis_mo1 | 無法核對 | The 200 MW is the first phase of the Blue Valley **power plant** (IPP generation), not a disclosed site or IT load. Other sources give that first plant phase as 250 MW. Nebius's chairman says the first data-center phase is operational Q2 2027, but gives no MW. | 0 | Under "How will the project be powered?": "The initial phase will deliver 200 MW … The power plant will be built in phases." | independencemo.gov FAQ; govtech 2026-01-28; kshb 2026-05-12; DCD |
| 23 | cp_nbis_pa1 | 屬實 | none (260 設施, 2027-10, Interested-party). Power is phased under the PPL Energy Service Agreement, which is utility load tied to this one campus. Replaces cpx_nbis_pa 1,200, which must stay at inc 0. | 1 | "260 MW available by October 2027, 660 MW by January 2029, and 1,000+ MW by November 2029" | nebius.com/schuylkillcountypa |
| 24 | cp_nbis_lpr1 | 屬實 (Derived, upper bound) | none (70 設施, 2027-12). The first-phase data-center MW is not disclosed, so the substation size is used as a proxy. | 1 | "In its first phase, the substation will have a capacity of 70-megawatt (MW)"; "expected to be completed in spring 2027" | caverion 2026-05-13 |
| 25 | cp_nbis_mts2 | 屬實 | none (145−75=70, 2027-12) | 1 | Dec 31 2027 145 projected: "Company estimates put completion of a second 70 MW expansion by year-end" | epoch nebius-mantsala |
| 26 | cp_orcl_shk1_edit | 屬實 | none (shk1 600 設施 @2027-11; shk2 1,400 @2028-11; 420/1,400 = 30% of 2,000). Optional: a finer split of 200 each @2027-08/10/11. | 1 | Aug 15 2027 140, Oct 19 2027 280, Nov 18 2027 420 … Nov 11 2028 1,400 MW IT (TDLR completion + 6 months) | epoch openai-stargate-shackelford |
| 27 | cpx_oai_lord_edit | 屬實 | none (214 IT, 2027-03, op oai). Epoch lists the owner as SoftBank and the user as OpenAI. | 1 (flip 0→1) | Mar 1 2027 projected 214 MW; derived from OpenAI's 1.5 GW for Lordstown+Milam minus Milam's 1.2 GW | epoch openai-stargate-lordstown |
| 28 | cp_neo_sines | 屬實 | none (233−33=200, 2027-09, neo/msft). Nscale owns the chips; Microsoft is the user. | 1 | Sep 30 2027 SIN02 233, "projected to be operational, based on the site factsheet" | epoch start-campus-sines-data-campus |
| 29 | cp_neo_lmar | 屬實 | none (378−210=168, 2027-03). Owner is AI XPV Platform, lessee Fluidstack, user likely Anthropic. No overlap with the twf (Kentucky), barber or other neo rows. | 1 | Mar 31 2027 Building 3 (CB5) 378; "Epoch expects about a three-month delay" | epoch anthropic-lake-mariner |
| 30 | cp_googl_vanburen | 小誤 | basis 未說明 → **設施** (the township FAQ says it will use roughly 1 GW of electricity). 1000, 2028-12, googl unchanged. MPSC approval 2026-10-01 confirmed; Google LLC is the named customer. | 1 | "expected to begin taking service in December 2027 with a maximum load achieved by December 2028" (MPSC 2026-10-01) | planetdetroit 2026-03-17; michigan.gov/mpsc 2026-10-01 |
| 31 | cp_amzn_kozani1 | 小誤 | basis 設施 → **未說明** (the source does not say IT or utility; 設施 only if TechCapital's "utility power" wording is checked). 300, 2028-12, amzn kept. AWS operates the IT; PPC supplies site, buildings and energy. The deal is a binding MoU. | 1 (MoU risk) | "initial capacity" 300 MW; "binding memorandum of understanding"; "The goal is to complete the project by 2028." | balkangreenenergynews 2026-09-18 |

## Specific questions

**(a) Goodnight**

- **Ownership.** Google's ownership or tenancy is supported by an openable source. W.media (2026-09-29) reports that Crusoe "has announced that it is the developer of Google's new data center campus." Energy Mag says state filings name no tenant, and cites Cleanview's report that Google confirmed working with Crusoe. Epoch's "speculative" label is more cautious than the current evidence; op=googl is reasonable. The "sold to Google" wording rests only on New Project Media, which comes from the gatherer notes and which I did not open.
- **Timing of the 2027-10 step.** The 6-buildings-by-2027-10 timing is grounded in filings:
  - a TDLR filing for Building 6 with completion 2027-07-30;
  - the original tax abatement deadline of EOY 2027;
  - an analogy with Abilene.
- **Size of the 2027-10 step.** The MW is not grounded. Epoch gives no per-building MW, and its 1,006 implies about 200 MW for each of Buildings 3–6, double Buildings 1–2. That conflicts with the abatement agreement ("four 25-megawatt data halls planned for each building", which is 100 MW per building) and with the first phase being "as much as 530 megawatts". The page cites no power or permit basis for the extra MW.
- **Correction.** 412 MW, which puts all six buildings at about 618 MW.

**(b) Red Oak**

- **Oracle link.** The tie to Oracle comes only second-hand: Dallas Innovates, citing Business Insider, which cited the Texas comptroller's tax-exemption list. DataBank and TDLR list the tenant as "Not Assigned".
- **MW.** 180 MW covers DFW9–11 together. Separately, the fourth building adds "60 megawatts of incremental IT", which supports 60 MW-IT per building.
- **Timing.** TDLR gives DFW10 tenant finish-out completion 2027-01-31 and DFW11 completion 2028-01-15. DCD gives DFW10 as Feb 2025 to Aug 2028, which is probably a different permit. Only DFW10 belongs in 2027.

**(c) Saline**

- **Ramp dates.** The DTE application of 2025-10-31 supports load starting January 2027 and maximum load by December 2027, with up to a 12-month delay.
- **The 1,383 figure.** It appears as the monthly committed figure in Exhibit A-1, with the note "8 Years Post Full Ramp Achieved (December 1, 2035)", which implies full ramp on 2027-12-01. The narrative text says "approximately 1.4 gigawatts".
- **The 1/3 split.** It is not in the source; the source gives no per-building or per-phase split. Oracle's 2026-06-01 release does confirm three equal 550,000 sq ft buildings, with the first "nearing completion".
- **Conflicting evidence.** Epoch's Stargate Michigan page puts all three buildings at 2028-12-31 (988 MW-IT), assuming the full delay. The contract itself would put all 1,383 at 2027-12. The 461 @2027-12 plus 922 @2028-12 split sits between the two, so it is reasonable and conservative.

**(d) Nebius**

- **Pennsylvania (Schuylkill).** The 260 MW figure and the October 2027 date are both stated verbatim on Nebius's page.
- **Lappeenranta.** The 70 MW first-phase substation is due in spring 2027 and the facility is "expected to be completed in 2027". The first-phase data-center MW itself is not disclosed.
- **Independence, Missouri.**
  - The 200 MW is the first phase of the power plant; GovTech and DCD say 250 MW by October 2027.
  - Nebius's chairman says the first data-center phase is operational in Q2 2027, but gives no MW.
  - The plant reopens in October 2027, with interim supply from June 2026.
  - Under the policy that generation capacity is not site capacity, this row cannot be verified as site MW.
  - If Andy decides that a plant dedicated to one campus counts like a utility contract, then 200 設施 @2027-10 Derived would be defensible.

**(e) AWS New Albany**

- Epoch counts only the share taken by leading AI chips: "which we estimate to be 44%".
- "The full IT power is much higher than 113 MW."
- So the 125 MW increment covers AI chips only. Full IT would be about 125/0.44 ≈ 284 MW.

## General findings

- **G1. Epoch rows.** All 21 Epoch-based rows (1–7, 9–13, 15–19, 25–29) match the pages exactly on both value and date. Every 2027 value is labelled *projected*; Epoch derives them from satellite imagery plus filing analogies, so the Analogy tag is right. Most user labels are "speculative" (Google DeepMind, Anthropic, Microsoft), but owner labels support the op assignments.
- **G2. Basis inconsistency for AWS.** AWS New Albany is AI-only (44%). Ridgeland, Berwick and New Carlisle are full Epoch IT, with no share applied. Mixing the two understates AWS general-purpose capacity. Decide one convention and annotate row 11.
- **G3. Base-stock dependency.** Several increments start from end-2026 Epoch points that are themselves *projected*:
  - Fort Wayne: 173 @2026-12-01;
  - Cedar Rapids: 87 @2026-12-31;
  - Lake Mariner: 210 @2026-10-31.

  These rows are correct only if the model's 2026 base already includes those projected points. If the base is observed-only, about 470 MW of 2026 capacity is missing, which could slip into 2027.
- **G4. Contract loads.** The contract-load rows (20, 23, 24, 30) are utility or substation capacity tied to one campus, so they are on a facility basis and consistent with policy. Row 22 is generation capacity and is not consistent with policy.
- **G5. Claim text.** Two claim texts mix up construction-completion dates (TDLR, Energy Mag) with operational dates: row 8 (Building 4 vs Building 1) and row 21 (the 2027-12 date was not supported by either TDLR end date). Make sure the claim column quotes the date actually used.
- **G6. Rows that replace cpx rows.** Rows 14, 22, 23, 24, 26 and 27 replace or split existing cpx or cp rows. When they are applied:
  - cpx_amzn_cip, cpx_nbis_pa and cpx_nbis_lpr must stay at inc 0, or be reduced to their residual;
  - cp_orcl_sal must drop to 922;
  - cp_orcl_shk1 and cp_orcl_shk2 must be edited, not added to.

  No double counts were found otherwise. Specifically, Lake Mariner does not overlap the neo rows, Goodnight matches nothing in T2, and Black Pearl Data Hall 4 is a subset of an inc-0 row.
- **G7. 2028 rows.** Rows 30 and 31 are 2028 rows. If this batch is meant to cover 2027 only, they are out of scope, but they are correct for the 2028 ledger.

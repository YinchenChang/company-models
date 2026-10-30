# v0.1.6 evidence pass — 2028-07 to 2032-12 (fetched 2026-10-31)

Files: `site_candidates.csv` (55 rows), `power_evidence.csv` (16 rows), `t2_export.csv` (T2 export from v0.1.5 sheet In, rows 184+), `scripts/build.py` (makes both CSVs and the summary).
MW is in the basis the source states. Summary IT figures convert 設施/未說明 ÷1.2 (that is an assumption for 未說明).

## Rule used when splitting
- If a source gives a ramp start and a "by year X", I split the MW evenly across year-ends from the start year to X. The even split is [Assumed] and the row is tagged [Derived].
- If a source gives no evidence of when service starts, I put the whole amount in year X (Pecos, Rome, Vizag, Aragón).
- Power contracts are kept only in power_evidence.csv. A site row that rests on a contract says so in `reason`.

## J rows marked 排 v0.1.6 (12): re-opened
| J# | Result |
|---|---|
| 5 OG&E–Google OK | KOSU confirms 1 GW over 5 yrs and 1.2 GW peak by 2036. "2031" comes only from EEI (2026-10-09). OCC approval still pending (OG&E release 2026-04-30). Split 200/yr 2027–31. inc=1 |
| 6 OK 2036 tail | Outside the window. inc=0 |
| 7 Google Nebraska | Confirmed: 1,000–3,000 MW, "as soon as 2029", from utility meeting documents. Google has not announced it, and Gage County held a moratorium hearing 2026-06-16. inc=0 (changed from the disposition) |
| 12 Vizag | Page has no MW and no online date. Only "gigawatt-scale" and $15B over 2026–2030. 1 GW @2030 [Derived]. inc=1 |
| 27 AWS N. Indiana | Amazon "2.4 gigawatts of data center capacity" confirmed. FactSet confirms the ramp from 2027-01 to end-2032 and the 1,200 MW reduction option by 2029-03. The IURC order (ord_46322_061726) gives the 15-yr term and 2027-01-01 energization, but the text I could read shows no year table. Firm 1,200 split 2027–32 → inc=1. Reducible 1,200 → inc=0 until the option lapses |
| 34 AWS×HUMAIN | 50 MW and 2028 confirmed. inc=0 because it likely sits inside the HUMAIN national totals already in T2 (cp_sov_hum2/hum3) |
| 35 PIT100 | Re-confirmed: no MW, phases 2030-04 to 2036-06. inc=0 |
| 36 La Porte 788 | **Not reproduced.** WebFetch of the SEC 10-Q (both URL forms) returns only text up to Note 3; Regulatory Matters and MD&A are cut off. The direct download was blocked by the session's permission policy (bot-check rule), and I did not try other workarounds. No secondary source has "788" or "2031": DCD failed to load, FourWeekMBA failed to load, Benzinga failed to load. What I could reach: 1,100 MW GenCo generation, service starts 2028, first building early 2029 (Baxtel 2026-10-19), and EEI lists 538 MW. inc=0 until Andy opens the 10-Q in a browser |
| 37 Pecos | Confirmed: ~2 GW over 5–7 yrs, behind-the-meter gas. The Microsoft page does **not** mention Chevron or a 2028 first-power date, so that part is unverified. 2,000 @2032. inc=1 |
| 38 Aragón | PIGA plan text confirmed ("4 años … primera fase", 10 years overall). The plan has **no MW**; the 50 MW per site comes from elDiario.es, which I did not re-open. 150 @2031. inc=1 |
| 39 Catawba | Confirmed. **Basis corrected to IT**: Aterio's "Low case (Critical IT, filing-supported) ~720 MW". Initial capacity Q4-27 to Q1-28, completion allowed through 2032. Split 144/yr. inc=1 |
| 55 Meta Lebanon | Still no MW for the first phase and no full-build year (Capital Chronicle, DCK). inc=0 |

## Epoch points dated 2028-07 or later: dedupe against T2
- Duplicates (inc=0): Stargate Michigan, New Mexico, Shackelford B8–10, Wisconsin, Milam (104 MW-IT left over, treated as noise), Amazon Ridgeland B6–14.
- New: CoreWeave Helios Phase III, 132 IT @2028-12 (T2 hel2 note said the timing was undisclosed). inc=1.
- New: YTL Johor, +171 IT @2028-08, user unnamed, booked as oth. inc=1 (drop it if oth is excluded).
- New: Stream Phoenix, +156 IT @2030-01, operator and user unknown. inc=0.

## New items found (not in the J list)
- Microsoft Rome GA: 422 MW, completion by 2032, from the DRI filing dated 2026-10-20 (via Baxtel/DCD). inc=1.
- OpenAI Effingham GA: 3.2 GW over 2028–2032. Georgia PSC approved the contract 2026-08-26. Up to 1,000 MW is flexible demand response. 800/yr 2029–32 (facility). inc=1. This is the largest new item.
- AWS Salem Twp (Susquehanna): Talen PPA of 1,920 MW reaches full volume "no later than 2032". After netting the AWS Berwick Epoch 169 IT (≈203 facility): 343/yr 2028–32. inc=1. **Check** that Epoch's "AWS Berwick" is this campus. Also note the PPA is energy supply, not metered load.
- Google Cheyenne: Black Hills 8-K says service starts late 2027 and reaches peak load in 2030. Only the 590 MW grid portion is proposed (inc=1). The 2.7 GW total resource mix suggests a peak load around 2.3 GW; the upside of ~1.7 GW is inc=0.
- Meta Hyperion 2→5 GW: Entergy's Q2 10-Q lists 7 CCCTs (5,278 MW) in service 2030–2031 for the Evest (Meta) second data center. LPSC decision is targeted for Dec-2026. A secondary source (GDH, 2026-07-20) says "full 5 GW around 2032". 1,500 @2031 + 1,500 @2032 proposed. **JUDGMENT: this overturns H1 (T2 cpx_meta_hyp5 is inc=0). Andy decides.**
- Campuses with MW but no date (inc=0): Google Monrovia 390 (IURC approved 2026-10-07), New Florence 1.2 GW, Project Mica 500, Hermantown 1,000 (EEI)/700 MW generation.

## Power evidence: certainty
- Approved by a regulator: Georgia Power–OpenAI (PSC 2026-08-26), DTE–Google (MPSC 2026-10-01), AES Indiana–Google (IURC 2026-10-07), WEC VLC tariff (PSCW order 2026-05-21; the deck also lists a Q4-2026 decision, which is inconsistent), NIPSCO–AWS (IURC order; the date 2026-06-17 is inferred from the file name).
- Pending: Entergy CCCTs (LPSC, Dec-2026), OG&E–Google (OCC), El Paso Electric–Meta (PUCT vote not yet scheduled), Black Hills CPGS siting permit, NiSource–Microsoft.
- Reduction options: NIPSCO–AWS can cut 1,200 MW by 2029-03-31; Georgia Power–OpenAI has up to 1,000 MW of flexible demand response.
- WEC Mount Pleasant 2.6 GW and Port Washington 1.3 GW (both "through 2030") are already covered by T2 (Fairwater WI 2,263 IT ≈ 2,716 facility; cp_orcl_pw 1,300). They are power evidence only, with no new MW.

## Searched, not found
- NiSource 10-Q "788 MW" passage (see above).
- Year-by-year ramp for NIPSCO–AWS (FactSet, IURC order text).
- Meta Lebanon MW for the first phase and a completion year.
- Online years for Google New Florence, Monrovia, Hermantown and Project Mica.
- AES Ohio–AWS "1,500 by 2031" (EEI only): the Dayton Daily News page failed to load and no primary source was found. Not added.
- Duke–AWS Richmond County and Entergy Mississippi–AWS Madison/Hinds: no MW or dates found.
- WEC–Microsoft ramp schedule beyond "2.6 GW through 2030".

## Not searched
- Dominion, AEP Ohio and Georgia Power aggregate large-load books. They are not tied to one campus (AEP: 63 GW incremental by 2030, 2026-05 8-K). The Dominion pages were not opened.
- Oracle sites beyond the five Stargate sites (no 2029+ Oracle-only campus turned up in one search). Microsoft Fairwater sites beyond WI/Atlanta. Meta sites other than Hyperion, Lebanon and El Paso. Non-US hyperscaler campuses after 2028.

## Quote caveat
The EEI list and a few other pages were summarised by the fetch tool, which would not return verbatim text. Those quotes are marked "paraphrase".

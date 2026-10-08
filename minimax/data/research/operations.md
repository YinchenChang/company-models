# MiniMax (稀宇科技, HKEX 00100): operating, pricing and compute data

Compiled 2026-10-08. Financial statements are in `financials.md` and are not repeated here.
Tags: **P** = company primary (platform.minimax.io docs, minimax.io blog/news, HKEX filings/results PR); **S** = media or third party; **IP** = company forward-looking claim or target. **UNVERIFIED** = recalled or single weak source, not confirmed. **D** = derived by me.

---

## 1. Model lineup and timeline (text/LLM, plus key multimodal)

| Model | Release | Total / active params | Context | Attention | Open weights / license | Tag | Source |
|---|---|---|---|---|---|---|---|
| MiniMax-Text-01 (+VL-01) | 2025-01-15 | 456B / 45.9B (MoE) | Up to 4M (inference); 1M on OpenRouter | Hybrid: 7 of every 8 layers Lightning (linear) attention, 1 softmax | Yes (GitHub/HF) | P | [minimax.io/news/minimax-01-series-2](https://www.minimax.io/news/minimax-01-series-2); [release notes](https://platform.minimax.io/docs/release-notes/models.md) |
| MiniMax-M1 (reasoning; 40k/80k thinking budgets) | 2025-06-17 (OpenRouter listing) | 456B / 45.9B | 1M | Hybrid Lightning attention + MoE | Yes | P | [arXiv 2506.13585](https://arxiv.org/abs/2506.13585); [OpenRouter](https://openrouter.ai/minimax) |
| MiniMax-M2 | 2025-10-27 | 229.9B / 9.8B (256 experts, top-8). Often quoted as "230B/10B" | 192K native (204,800 on API/OpenRouter) | **Full attention** (MiniMax moved away from linear attention; blog "Why did M2 end up as a full attention model") | Yes, modified-MIT | P | [M2-series tech report arXiv 2605.26494](https://arxiv.org/abs/2605.26494); [blog](https://www.minimax.io/news/why-did-m2-end-up-as-a-full-attention-model); [Simon Willison 2025-10-29](https://feeds.simonwillison.net/2025/Oct/29/minimax-m2/) |
| MiniMax-M2.1 (+highspeed) | 2025-12-22 | 230B / 10B (same base family) | 204.8K | Full | Yes | P/S | [release notes](https://platform.minimax.io/docs/release-notes/models.md); [OpenRouter](https://openrouter.ai/minimax) |
| MiniMax-M2-her (powers Talkie/Xingye) | Q4 2025 (OpenRouter listing 2026-01-23) | not found | 66K | not found | not found | P | [FY2025 results](https://www.minimax.io/news/minimax-global-announces-full-year-2025-financial-results) |
| MiniMax-M2.5 (+Lightning/highspeed) | 2026-02-12 | 230B / 10B | 196–205K | Full | Yes, modified-MIT | P/S | [release notes](https://platform.minimax.io/docs/release-notes/models.md); [WorkOS](https://workos.com/blog/minimax-m25-most-popular-model-openrouter) |
| MiniMax-M2.7 (+highspeed) | 2026-03-18; weights about 2026-04-10 | 230B / 10B | 204.8K | Full | Weights on HF, but under a **non-commercial** modified-MIT license (commercial use needs authorization) | P/S | [release notes](https://platform.minimax.io/docs/release-notes/models.md); [Artificial Analysis on X](https://x.com/ArtificialAnlys/status/2044265942021255428); [letsdatascience 2026-04-13](https://letsdatascience.com/news/minimax-revises-license-after-releasing-m27-weights-04b47c74) |
| **MiniMax-M3** | 2026-06-01 (weights on HF about 06-07; tech report arXiv 2606.13392 on 06-11) | **~428B / ~23B** (427.04B exact; 60 layers; 128 routed experts top-4 + 1 shared; 7 MTP modules). Natively multimodal (text, image and video in) | 1M (1,048,576); 512K guaranteed on API | **MSA (MiniMax Sparse Attention)**: block-sparse on a GQA backbone, top-16 of 128-token KV blocks, dense in the first 3 layers | Yes, "MiniMax Community License": attribution required, and revenue above US$20M needs authorization | P/S | [M3 blog](https://www.minimax.io/blog/minimax-m3); [GitHub](https://github.com/MiniMax-AI/MiniMax-M3); [SemiAnalysis InferenceX](https://inferencex.semianalysis.com/model/minimax-m3) |
| M3 efficiency claim | 2026-06-01 | — | — | Per-token compute at 1M context is 1/20 of the previous generation; prefill >9x faster, decode >15x faster than M2 | — | IP/P | [M3 blog](https://www.minimax.io/blog/minimax-m3) |
| M3.1 | Announced at WAIC 2026-07-17. **M3.1-Flash-Preview** released 2026-09-27, only inside MiniMax Code / M Plan (no standalone API) | Not published | 1M | Not published | No weights or model card yet | P/S | [DataNorth 2026-09-28](https://datanorth.ai/news/minimax-releases-m3-1-flash-preview); [M Plan docs](https://platform.minimax.io/docs/m-plan/intro.md) |
| M3.1 cost target | 2026-08-26 interim call | Inference cost about 1/3 of M3 at launch (a 2/3 cut). Part of the saving goes to price/volume, part to gross margin; text expected to drive H2-26 GM improvement | — | — | — | IP | [QQ News 2026-08-27](https://news.qq.com/rain/a/20260827A08KN500); [locdd, citing The Paper](https://www.locdd.com/t/topic/86115) |
| M3 Pro (planned) | Not dated | About **3T** params | — | MSA 2.0 (smaller KV cache, higher prefill-cache hit rate, ~3x compute efficiency at the top size) | — | IP | Same as above |
| 10T model (long-term) | — | Path is 3T then 10T; about 200T tokens of data needed | — | — | — | IP | [QQ interview 2026-06-16](https://news.qq.com/rain/a/20260616A04QLX00) |
| Hailuo-02 / 2.3 / 2.3-Fast (video) | 2025-06-18 / 2025-10-28 | — | — | — | Closed | P | [release notes](https://platform.minimax.io/docs/release-notes/models.md) |
| MiniMax H3 / H3-Max (video) | 2026-07-31 / ~2026-09-02 | — | — | — | H3 open weights: >24M downloads in 3 weeks, >300 derivatives | P/S | [release notes](https://platform.minimax.io/docs/release-notes/models.md); [Huxiu](https://www.huxiu.com/article/4886381.html) |
| Speech-02 / 2.5 / 2.6 / 2.8 | 2025-04-02 / 2025-08-06 / 2025-10-29 / 2026-01-23 | — | — | — | Closed | P | [release notes](https://platform.minimax.io/docs/release-notes/models.md) |

Other model facts:
- M2 pre-training used 29.2T tokens (P, [arXiv 2605.26494](https://arxiv.org/abs/2605.26494)).
- M2 to M2.5 took 108 days (P, [KrASIA, FY25 call](https://kr-asia.com/minimaxs-arr-tops-usd-150-million-as-it-pivots-toward-an-ai-platform-model)).
- M3 decode speed was about 20 TPS at launch and about 70 TPS by mid-June, with a target of 80 (P/IP, CEO, [QQ 2026-06-16](https://news.qq.com/rain/a/20260616A04QLX00)).
- Text-model compute throughput per unit tripled in the roughly two months to late August 2026 (IP/P, interim call, [QQ 2026-08-27](https://news.qq.com/rain/a/20260827A08KN500)).

---

## 2. Pricing

### 2a. Text API list prices (USD per 1M tokens, international platform)

| Model | Period | Input | Output | Cache read | Cache write | Tag | Source |
|---|---|---|---|---|---|---|---|
| MiniMax-Text-01 | From 2025-01-15 | 0.20 | 1.10 | — | — | P | [minimax-01 news](https://www.minimax.io/news/minimax-01-series-2) |
| MiniMax-M1 | From 2025-06 | 0.40 (≤200K) | 2.20 | — | — | S | [OpenRouter](https://openrouter.ai/minimax); [pricepertoken](https://pricepertoken.com/pricing-page/model/minimax-minimax-m1) |
| M1 official CNY tiers | 2025-06 | ¥0.8 / ¥1.2 / ¥2.4 (0–32K / 32–128K / 128K–1M) | ¥8 / ¥16 / ¥24 | — | — | **UNVERIFIED** (recall) | — |
| MiniMax-M2 | 2025-10-27 (free until 2025-11-07) | 0.30 | 1.20 | 0.03 | 0.375 | P | [Willison](https://feeds.simonwillison.net/2025/Oct/29/minimax-m2/); [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| M2.1 / M2.5 / M2.7 standard | Launch to now | 0.30 | 1.20 | 0.03 (M2.x); 0.06 (M2.7) | 0.375 | P | [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| M2.x / M2.7 highspeed ("Lightning") | Launch to now | 0.60 | 2.40 | 0.03 / 0.06 | 0.375 | P | Same |
| M2.5-Lightning (third-party quote) | 2026-02 | 0.30 | 2.40 | — | — | S (conflicts with the official 0.60 input) | [WorkOS](https://workos.com/blog/minimax-m25-most-popular-model-openrouter); [Verdent](https://www.verdent.ai/guides/minimax-m2-5-pricing) |
| CNY equivalents (M2.1/M2.7) | 2025-12 to 2026 | ¥2.1 | ¥8.4 | — | — | S | [Soochow (东吴) IPO report](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf); [codepick](https://codepick.dev/en/guides/minimax-token-plan/) |
| **MiniMax-M3 launch list** | 2026-06-01 to 06-14 | **0.60** (≤512K) | **2.40** | — | — | S | [KuCoin 2026-06-18](https://www.kucoin.com/news/flash/minimax-cuts-m3-model-prices-by-50-permanently-after-developer-backlash); [aireiter](https://aireiter.com/blog/minimax-m3-1-release-api) |
| **MiniMax-M3 after the permanent 50% cut** | From **2026-06-15** | **0.30** (≤512K); 0.60 (>512K) | **1.20**; 2.40 | 0.06; 0.12 | not listed | P | [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| M3 Priority tier | 2026 | 0.45 / 0.90 | 1.80 / 3.60 | 0.09 / 0.18 | — | P | Same (1.5x standard) |
| M3 on OpenRouter (CoreWeave base) | 2026-10 | 0.23 | 0.96 | 0.05 | — | S | [OpenRouter M3](https://openrouter.ai/minimax/minimax-m3) |
| M3 CNY (domestic) | 2026-06 | ¥2.1 (≤512K) | — | — | — | S | [Sohu](https://www.sohu.com/a/1036968995_122496371) |
| M3 blended realized price | 2026-08 | US$0.22/M (from the caller's brief) | | | | Primary source **not found** in this search; treat as UNVERIFIED | — |
| M3.1-Flash-Preview | 2026-09-27 | No per-token price. Available only through Token/M Plan subscription | | | | P/S | [DataNorth](https://datanorth.ai/news/minimax-releases-m3-1-flash-preview) |

Price history in short:
- The list price of $0.30/$1.20 has held flat for the whole M2 family from Oct-2025 to now.
- M3 launched at 2x that ($0.60/$2.40) on 2026-06-01. It was cut back to $0.30/$1.20 on 2026-06-15 after developer backlash, so at the same price M3 gives more capability.
- The interim results said "M3 pricing unchanged vs M2 series" (S, [cs.com.cn](https://jnzstatic.cs.com.cn/zzb/htmlInfo/131401.html)).
- A third-party tracker reported a top-tier per-token rate rising from $3 to $4.80 on 2026-08-08 (S, [costbench](https://costbench.com/changelog/minimax-api-price-increase-2026-08/)). It does not say which product this covers, so it is low confidence.

### 2b. Subscriptions for developers and agents (Coding Plan → Token Plan → M Plan)

| Plan | Period | Tiers and prices | Quotas | Tag | Source |
|---|---|---|---|---|---|
| Coding Plan (USD, per prompt, M2/M2.1, text only) | ~Oct-2025 to ~Mar-2026 | Starter $10, Plus $20, Max $50 per month (annual $100/$200/$500) | 100 / 300 / 1,000 prompts per 5h | S | [Verdent 2026-02-23](https://www.verdent.ai/guides/minimax-m2-5-pricing) |
| Token Plan v1 (CNY, per request, M2.7, all modalities) | ~2026-03-22 launch | Starter ¥29, Plus ¥49, Max ¥119; Highspeed Plus ¥98, Max ¥199, Ultra ¥899 | 600 / 1,500 / 4,500 / 30,000 requests per 5h | S | [codepick](https://codepick.dev/en/guides/minimax-token-plan/); [Sohu](https://www.sohu.com/a/1036968995_122496371) |
| Token Plan v2 (per-token billing with M3) | 2026-06-01 | Plus $20 (~1.7B tokens/mo), Max $50 (~5.1B), Ultra $120 (~9.8B). CNY Starter went from ¥29 to ¥49 | 5h rolling + weekly | P | [M3 blog](https://www.minimax.io/blog/minimax-m3) |
| Backlash | 2026-06-01 to 06-05 | Users measured effective cost increases of up to 257% (example: 3–5B tokens/mo went from ¥49 to ¥175). Plus went from ~1,500 to 300–500 calls per 5h. Apology on 06-02, with compensation (quota reset; +50% weekly quota for subscribers from 3/22 to 6/5) | | S | [Sohu](https://www.sohu.com/a/1036968995_122496371); [KuCoin](https://www.kucoin.com/news/flash/minimax-cuts-m3-model-prices-by-50-permanently-after-developer-backlash) |
| Token Plan price rise | ~2026-08-26 | Plus $22, Max $55, Ultra $132 (+10%). Quotas unchanged | | S, current prices confirmed P | [usagepricing](https://usagepricing.com/blueprint/activity/minimax-2026-08-26-price-change); [token plan page](https://platform.minimax.io/docs/guides/pricing-token-plan.md) |
| Credits (overflow) | 2026 | 1,000 credits = $1 ($5 / $25 / $100 packs, valid 365 days) | | P | [token plan page](https://platform.minimax.io/docs/guides/pricing-token-plan.md) |
| **M Plan** (replaces Token Plan) | Announced 2026-09-29 | Go $22, Explore $55, Build $132 (first month 50% off until 10-14) | Explore = 3x Go, Build = 7.5x Go. Includes M3.1-Flash-Preview; H3 video only in Explore and Build | P | [announcement](https://platform.minimax.io/docs/token-plan/announcements.md); [monthly offer](https://platform.minimax.io/docs/m-plan/monthly-offer.md); [M Plan intro](https://platform.minimax.io/docs/m-plan/intro.md) |
| Token Plan for Teams | 2026 | Page exists; prices not retrieved | | P | [link](https://platform.minimax.io/docs/guides/pricing-token-plan-team.md) |

Implied subscription token price (D): $20 for 1.7B tokens is about **$0.012/M**, and $120 for 9.8B is about **$0.012/M**, if quotas are fully used. That is roughly 1/20 of the blended API price, so subscriptions are a much cheaper channel when heavily used.

### 2c. Video API (Hailuo / H3)

| Model | Unit price | Tag | Source |
|---|---|---|---|
| Hailuo-02 | 512P 6s $0.10; 768P 6s $0.28 / 10s $0.56; 1080P 6s $0.49 | P | [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| Hailuo-2.3 | 768P 6s $0.28 / 10s $0.56; 1080P 6s $0.49 | P | Same |
| Hailuo-2.3-Fast | 768P 6s $0.19 / 10s $0.32; 1080P 6s $0.33 | P | Same |
| Hailuo-2.3 CNY | 768P 6s ¥2.00; 1080P 6s ¥3.50; Fast I2V 768P 6s ¥1.35 | S | [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf) |
| MiniMax-H3 | 768P $0.08/s; 2K $0.13/s (a 6s 768P clip is about $0.48) | P | [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| MiniMax-H3-Max | 480P $0.05/s; 768P $0.08/s; regeneration 768P→2K $0.05/s | P | Same |
| Video packages | Separate page, not retrieved | P | [pricing-video](https://platform.minimax.io/docs/guides/pricing-video.md) |

### 2d. Speech, audio, image and music API

| Item | Price | Tag | Source |
|---|---|---|---|
| speech-2.8-hd / 2.6-hd / 02-hd | $100 per M characters ($1.00 per 10K) | P | [paygo page](https://platform.minimax.io/docs/guides/pricing-paygo.md) |
| speech-2.8-turbo / 2.6-turbo / 02-turbo | $60 per M characters ($0.60 per 10K) | P | Same |
| CNY: speech-2.6-hd / turbo | ¥3.5 / ¥2.0 per 10K characters | S | [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf) |
| ASR | $0.38/hour | P | paygo page |
| Voice design / rapid clone | $3 / $1.50 per voice | P | paygo page |
| Audio subscription | Starter $5, Standard $30, Pro $99, Scale $249, Business $999 per month (100K to 20M audio points) | P | [pricing](https://platform.minimax.io/docs/guides/pricing) |
| image-01 | $0.0035/image (CNY ¥0.025) | P | paygo page |
| Music-2.6/3.0 | $0.15 per song (≤5 min); paid music APIs closed to new users from 2026-08-20 | P | paygo page |

---

## 3. Token volumes

| Metric | Value | Date | Tag | Source |
|---|---|---|---|---|
| Open platform daily tokens (prospectus era) | ">1 trillion tokens/day" | ~2025-09/12 | S (QbitAI summary of prospectus; not checked against the filing) | [QbitAI 2025-12](https://www.qbitai.com/2025/12/363445.html) |
| M2 on OpenRouter | First Chinese model above 50B tokens/day | Late 2025 | P | [FY2025 results](https://www.minimax.io/news/minimax-global-announces-full-year-2025-financial-results) |
| M2-series daily tokens | Feb-26 more than 6x Dec-25 | 2026-02 | P | Same |
| Coding Plan tokens | Feb-26 more than 10x Dec-25 | 2026-02 | P | Same |
| Total token consumption | **Jul-26 = 20x Jan-26** | 2026-07 | P | [1H26 results](https://www.minimax.io/news/minimax-announces-first-half-2026-financial-results-1787744160) |
| Caveat on the 20x | Partly inflated by the switch from per-request to per-token billing and by the 50% price cut | 2026-08 | S | [Huxiu](https://www.huxiu.com/article/4886381.html) |
| 2026 outlook | Token volume could grow 1–2 orders of magnitude in 2026 | 2026-03 | IP | [KrASIA](https://kr-asia.com/minimaxs-arr-tops-usd-150-million-as-it-pivots-toward-an-ai-platform-model) |
| CEO anecdote | M2 target was "about 100M (一亿) tokens/day". M2.7 ran about 10x that target. M3 "exceeded expectations" | 2026-06 | P (quote; units ambiguous) | [QQ 2026-06-16](https://news.qq.com/rain/a/20260616A04QLX00) |
| Absolute company-wide daily tokens in 2026 | **not found** (not disclosed) | | | |

### OpenRouter (S)

| Date | MiniMax data point | Source |
|---|---|---|
| 2026-02 | M2.5 was #1 model. 2.45T tokens in one week ([WorkOS](https://workos.com/blog/minimax-m25-most-popular-model-openrouter)); 3.07T per week ([Pandaily](https://pandaily.com/chinese-models-top-open-router-token-rankings-as-agent-scenarios-emerge-as-new-frontier)); about 4.55T in the month ([presenc](https://presenc.ai/research/openrouter-model-usage-rankings-2026)) | as listed |
| 2026-04-01 | M2.7 #3 model at 1.92T per week. MiniMax 9.2% of weekly tokens as an author | [digitalapplied](https://www.digitalapplied.com/blog/openrouter-rankings-april-2026-top-ai-models-data) |
| 4 weeks to 2026-04-14 | MiniMax 7.24T per week, **9.5% share** (up from 0% a year earlier) | [codesota](https://www.codesota.com/agentic/openrouter-trends) |
| 2026-06 | MiniMax #6 author, 2.37T per week, 8.1% share. M3 #3 by daily tokens (447B/day) | [nodemini](https://nodemini.com/en/blog/2026-openrouter-rankings-june-chinese-models-61-percent.html) |
| 2026-08-05 | M3 still in the top 10 weekly (no number given) | [KuCoin](https://www.kucoin.com/news/flash/chinese-models-dominate-openrouter-weekly-token-usage-ranking) |
| 2026-10 (trailing 30 days) | **MiniMax not in the top 10 models.** The #10 model has 15.7T per month. Not in the top-9 authors by request share (week of 09-28). The OpenRouter M3 page shows 1.28T tokens for M3 (period unclear) | [OpenRouter rankings](https://openrouter.ai/rankings?view=month); [OpenRouter MiniMax](https://openrouter.ai/minimax) |
| Aug-13 to Oct-7, 2026 (OpenCode telemetry) | M3 about 2.7T tokens over the period (down 71%). Daily tokens fell from a 107B peak (Aug-17) to 20–32B (late Sep to early Oct). #18 model, 94% of input cached | [opencode.ai](https://opencode.ai/data/minimax/minimax-m3) |

Read-across (D): MiniMax's third-party share peaked around Feb–Apr 2026 at about 9–10% of OpenRouter tokens and has faded since mid-2026. DeepSeek V4 Flash, GLM 5.3 Flash, Tencent Hy4, MiMo and others overtook it. Company-reported growth (20x by July) is therefore mainly first-party, from the Token Plan and direct enterprise API.

### ARR series (company metric: weekly revenue x 52)

| Date | ARR | Tag | Source |
|---|---|---|---|
| 2026-02 | >US$150M | P | [KrASIA](https://kr-asia.com/minimaxs-arr-tops-usd-150-million-as-it-pivots-toward-an-ai-platform-model) |
| 2026-05 | ~US$400M; monthly growth >25% | S | [Huxiu](https://www.huxiu.com/article/4886381.html) |
| 2026-08 | >US$800M; ~80% B2B / 20% B2C (a year earlier ~30/70) | P | [QQ](https://news.qq.com/rain/a/20260828A0AX5J00) |
| FY2026 target | US$1B ARR (not raised) | IP/S | [Huxiu](https://www.huxiu.com/article/4886381.html) |

ARR method: Huxiu says weekly revenue x 52. One QQ article says "x 365" (likely meaning daily x 365). Either way it is a run-rate on a peak week.

Implied volume cross-check (D, rough): B2B ARR of about $640M at a $0.22/M blend would be about 2.9 quadrillion tokens a year, or about 8T tokens a day. This is an upper bound. B2B ARR also includes Token Plan subscriptions (about 80% enterprise-sourced per Huxiu), whose implied price per token is far lower, and non-text APIs. Treat it only as order of magnitude.

---

## 4. Consumer products

| Metric | Value | Date | Tag | Source |
|---|---|---|---|---|
| Cumulative individual users | >212M (Talkie/Xingye 147M; Hailuo 42.3M; MiniMax app 19.1M) | 2025-09-30 | P (prospectus, via media) | [QQ prospectus summary](https://news.qq.com/rain/a/20251222A01Z4300) |
| Cumulative users | >236M (FY25); >300M (1H26) | 2025-12 / 2026-06 | P | [FY25 PR](https://www.minimax.io/news/minimax-global-announces-full-year-2025-financial-results); [1H26 PR](https://www.minimax.io/news/minimax-announces-first-half-2026-financial-results-1787744160) |
| Paying users (AI-native products) | 119.7K (2023), 650.3K (2024), 1,771.6K (9M25) | | P (prospectus, via media) | [QQ](https://news.qq.com/rain/a/20251222A01Z4300); [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf) |
| Open-platform paying customers (≥$50 spend) | ~100 (2023), ~700 (2024), ~2,500 (9M25) | | P (via media) | Same |
| Talkie/Xingye MAU | ~20M (9M25 average); marketing cut 90% | 2025 | P (via ChinaTalk) | [ChinaTalk](https://www.chinatalk.media/p/zhipu-and-minimax-ipo) |
| Hailuo AI MAU | ~5.6M (9M25) | 2025 | P (via ChinaTalk) | Same |
| Talkie average spend | ~US$5 per customer (9M25), down from 2024 | 2025 | S | Same |
| Talkie app MAU (third party) | **32.85M** (#16 globally, +12.6% MoM) | 2026-06 | S | [aicpb](https://www.aicpb.com/en/ai-rankings/products/global-ai-rankings/apps) |
| Xingye (星野) app MAU | 3.71M (−3.2% MoM) | 2026-06 | S | [aicpb China](https://www.aicpb.com/ai-rankings/products/china-ai-rankings/apps) |
| MiniMax Agent app MAU | 3.79M (−1.3% MoM) | 2026-06 | S | Same |
| Hailuo MAU, 2026 | **not found** (not in aicpb app top 100) | | | |
| Talkie revenue (9M25) | US$18.75M (35.1% of revenue); Hailuo ~32.6%; open platform 28.9% | 2025 | P (via media) | [QQ](https://news.qq.com/rain/a/20251222A01Z4300) |
| Talkie subscription | Talkie+ ~$9.99/month; tiers up to $199.99/month | 2026 | S | [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf); [aipicks.jp](https://aipicks.jp/mag/talkie-ai-guide-2026) |
| Hailuo AI web plans | Standard $7.99/mo (1,000 credits); Pro $24.99 (4,500); Master $63.99 (10,000); Max $199.99 (20,000) | Verified by source 2026-07-01 | S | [costbench](https://costbench.com/software/ai-video-generators/hailuo-ai) |
| MiniMax Agent plans and prices | **not found** (now folded into Token/M Plan credits) | | | |
| Videos generated | >590M (9M25); >600M (FY25) | | P | FY25 PR |
| Speech generated | >200M hours (FY25) | | P | FY25 PR |
| C-end daily usage | >70 min per user per day | 2025 | S | Soochow |
| Paying users / MAU in 1H26 | **not disclosed** | | | 1H26 PR |
| AI-native revenue 1H26 | US$42.6M (+100.9%) | | P | 1H26 PR |
| Enterprise and developer base | 214K (FY25); >1M (1H26 PR); ">2M, ~10x end-2025" (August call) | | P | PRs; [QQ](https://news.qq.com/rain/a/20260828A0AX5J00) |

---

## 5. Compute

| Item | Data | Date | Tag | Source |
|---|---|---|---|---|
| M1 RL training | 512 H800 for about 3 weeks; rental cost **US$534.7K** | 2025-06 | P | [arXiv 2506.13585](https://arxiv.org/abs/2506.13585) |
| M2/M2.5/M2.7 training cost | **not found** (tech report gives 29.2T pre-training tokens, no GPU-hours) | | P | [arXiv 2605.26494](https://arxiv.org/abs/2605.26494) |
| M3 training compute | **not found**. The MSA report benchmarks on H800. M3 kernel case study cites Hopper FP8 | | P/S | [InferenceX](https://inferencex.semianalysis.com/model/minimax-m3) |
| Asset model (at IPO) | "Light-asset": no owned training cluster. Multiple cloud vendors (5 Chinese, 2 Singapore, 1 both) | 2025 | P (via ChinaTalk) | [ChinaTalk](https://www.chinatalk.media/p/zhipu-and-minimax-ipo) |
| Compute model (2026) | Self-controlled core cluster for flagship training; cloud vendors for elastic capacity; "Token Factory" for extra inference. 97% effective training time; current and planned compute supports a 3T model; text gets about 4x the training resources of video | 2026-08 | P/IP | [QQ 2026-08-27](https://news.qq.com/rain/a/20260827A08KN500) |
| Training-related cloud spend | $4.15M (2022), $47.2M (2023), ~$140M (2024), $142M (9M25; 79% of R&D; 267% of revenue) | | P (prospectus, via media) | [QbitAI](https://www.qbitai.com/2025/12/363445.html); [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf) |
| R&D (mostly cloud training) | $252.8M FY25; $296.9M 1H26 (+138.8%) | | P | PRs |
| **Alibaba Cloud procurement caps** (continuing connected transaction) | Revised Aug-2026 to RMB ~2.1B / 2.9B / 3.6B for 2026/27/28 (≈ US$300M / 400M / 500M). Total ~RMB 8.6B (+220%). Prior caps were RMB 0.8B / 0.9B / 1.0B (≈ $115M for 2026). **1H26 usage was 65.7% of the old 2026 cap** (≈ RMB 525M / ~$75M in 6 months). For training and inference | 2026-08-26/27 | P (HKEX announcement, via media) | [TechNode](https://cn.technode.com/post/2026-08-31/minimax-alibaba-cloud-spending/); [QQ](https://news.qq.com/rain/a/20260827A04F7D00) |
| API services sold to Alibaba (cap) | 2026 cap raised from RMB 4.65M to RMB 54M (~$7.5M); ~$33M by 2028 | 2026-08 | P (via media) | Same |
| Alibaba as customer | ~22% of 2024 revenue | 2024 | S (single source, verify) | [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf) |
| Other compute supplier | 鸿博股份 / 英博数科 dedicated compute named as a partner | 2025 | S, **UNVERIFIED** (no contract value found) | Soochow |
| Domestic chips | M3 and H3 being adapted to domestic chips. A "large-scale domestic compute cluster coming online soon" will take production traffic. CEO mentioned Huawei (Ascend "990"?) in June. **No vendor, card count or MW disclosed** | 2026-06 / 08 | P/IP | [QQ 2026-08-26](https://news.qq.com/rain/a/20260826A0D8YS00); [QQ 2026-06-16](https://news.qq.com/rain/a/20260616A04QLX00) |
| July 2026 capital raise | Placement of HK$9.54B at HK$268 plus HK$6.5B convertible bonds (conversion price HK$335), about HK$16B / US$2B in total. About 80% earmarked for AI infrastructure and model R&D | 2026-07-10 | P (via media) | [10jqka](https://stock.10jqka.com.cn/20260710/c678088351.shtml); [QQ](https://news.qq.com/rain/a/20260827A04F7D00) |
| Gross margin by business | B2B (open platform) 69.4% vs C-end 4.7% (9M25); 62% B2B in 2024. Group GM 12.2% (2024) → 25.4% (FY25) → 17.9% (1H26). H2-25 was 33.5%. Management expects only sequential H2-26 improvement, not a return to 33.5% | | P (prospectus/results), S (half-year split) | [Soochow](https://pdf.dfcfw.com/pdf/H3_AP202601081816847687_1.pdf?1767890362000.pdf); [Huxiu](https://www.huxiu.com/article/4886381.html) |
| Inference cost per token | **Not disclosed in absolute terms.** Only relative claims: M3 at 1M context uses 1/20 the compute per token of the previous generation; M3.1 targets 1/3 of M3's launch cost; unit throughput 3x in about 2 months; MSA 2.0 ~3x | | IP | Above |
| GPU count, cluster size or MW (company or analyst) | **not found**. No analyst estimate of MiniMax's total GPUs or MW was found in public sources | | | |
| Third-party serving | M3 served on OpenRouter by 13 providers (CoreWeave base listing, GMICloud, DeepInfra, Together, SambaNova, etc.). M2.5 runs on AMD hardware at some hosts. NVIDIA blog on M2.7 | 2026 | S | [OpenRouter](https://openrouter.ai/minimax/minimax-m3); [NVIDIA](https://developer.nvidia.com/blog/minimax-m2-7-advances-scalable-agentic-workflows-on-nvidia-platforms-for-complex-ai-applications/) |

Rough compute sizing (D, for the model only): Alibaba Cloud alone at about $75M in 1H26 annualizes to about $150M. The new 2026 cap is $300M. At a blended ~$2/GPU-hour for H800/H20-class rental, $300M a year is about 17K GPU-equivalents running continuously. That excludes other vendors and the self-controlled cluster. This is my estimate only.

---

## 6. Chinese LLM API pricing context (2026)

| Provider | Data point | Date | Tag | Source |
|---|---|---|---|---|
| DeepSeek | V4 launched at a 75% promotional discount (Apr). Made permanent on 05-22/23: V4-Pro $0.435 input / $0.87 output; V4-Flash $0.14 / $0.28 | 2026-04/05 | S | [apifox](https://apifox.com/apiskills/2026-chinese-llm-price-war-api-costs-comparison-4/) |
| DeepSeek | **Price increase from 2026-08-16** with peak and off-peak tiers. V4-Pro peak $1.32 / $3.96 (off-peak $0.66 / $1.98). V4-Flash peak $0.44 / $1.32 (off-peak $0.22 / $0.66). Cache-hit up to +1,100%. Reason given: capacity strain | 2026-08 | S | [TechTimes](https://www.techtimes.com/articles/324764/20260817/deepseek-v4-api-prices-quadruple-peak-what-developers-pay-starting-now.htm) |
| DeepSeek V4.1 Flash | Peak $0.30 / $1.20; off-peak $0.15 / $0.60; cache $0.003–0.006 | 2026-10 | S | [devtk.ai](https://devtk.ai/en/blog/chinese-ai-models-api-pricing-2026/) |
| Zhipu GLM | GLM-5 $1.00 / $3.20 (about +30% vs GLM-4.7 at launch); GLM-5.1 $0.98 / $3.08; GLM-5.3 $1.40 / $4.40; GLM-5.3 Flash $0.15 / $0.50 | 2026 | S | apifox; devtk.ai |
| Kimi (Moonshot) | K2.6 tiered input $0.16–2.00, output about $2.50. **Kimi K3 (2026-07-16): $3.00 / $15.00**, cache $0.30 | 2026 | S | Same |
| Qwen (Alibaba) | Qwen3 Max $0.78 / $3.90. Qwen3.8 Max ¥12 / ¥36; Qwen3.8 Flash ¥0.8 / ¥2.7 | 2026 | S | Same |
| Xiaomi MiMo | MiMo-2.6 Flash $0.14 / $0.28; Pro $0.435 / $0.87 | 2026-10 | S | devtk.ai |
| Overall | Six API price cuts by major Chinese labs in 1H26 (three made permanent). From 2H26 the trend reversed toward **increases**: Kimi, Zhipu and MiniMax subscription billing changes, DeepSeek's peak pricing, and MiniMax Token Plan +10%. Goldman projects China's daily token use at 350T by end-2026 | 2026 | S | apifox; [Sohu](https://www.sohu.com/a/1036968995_122496371) |

Implication (D): the "prices fall 50–80% a year" assumption held through 1H26. Since then the frontier price in China has been flat to rising, driven by capacity constraints. MiniMax's API list price has been flat at $0.30/$1.20 since Oct-2025. Its effective price fell mainly through the mix shift to subscriptions and caching.

---

## Gaps (not found)
- Absolute company-wide daily or weekly token volume for 2026.
- Primary source for the M3 blended price of $0.22/M.
- GPU types and counts, cluster MW, domestic-chip vendor and cluster size. No analyst GPU or MW estimates found.
- Full prospectus list of cloud suppliers and spend by vendor. The 鸿博 contract was not verified.
- Training cost for M2.x and M3.
- 2026 MAU and paying users from the company. Hailuo MAU for 2026.
- MiniMax Agent standalone plan prices.
- M1 official tiered CNY price (only recalled, UNVERIFIED).

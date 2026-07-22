# CNTB Deep-Dive — Full Project Context (Handoff Document)

**Purpose of this document:** everything needed to pick this project back up in a fresh session/tool — Claude Code, another Claude chat, or a human analyst. It covers the original brief, all research findings, every assumption and its source, all corrections made against the primary 10-Q filing, the exact file/row structure of the Excel model (so it can be edited programmatically without breaking formulas), and the list of open items still worth chasing.

---

## 1. Original brief

The user is running an ongoing biotech investment-analysis working file on **Connect Biopharma Holdings Limited (NASDAQ: CNTB)**, lead asset **rademikibart** (anti-IL-4Rα monoclonal antibody), for **acute exacerbations of asthma and COPD**. The brief (from `Biotech_DD_Prompt_Template.md`, the user's own reusable template) asks for:

1. A fully formula-driven **Excel valuation model** (no hardcoded outputs) with: Cover & Assumptions (Low/Base/High + single scenario selector flowing through every tab), Balance Sheet & Cash Runway, rNPV Model (bottom-up, epi-driven, most conservative number), DCF Model (unrisked ceiling case, explicit terminal value), Relative Valuation (genuine peers, EV/peak-sales, not a share-count artifact), Price Target Summary (blended, editable weights, sensitivity).
2. A **Word thesis memo**: Executive Summary, Long Thesis, Origin Story, Key Risks (regulatory/clinical, commercial, financial, ownership/governance), Balance Sheet Summary, Ownership & Insider Activity, Where Outside Expertise Would Help, Valuation Summary Table.
3. **Sell-side & third-party research reconciliation** (analyst targets, conflicts of interest, named public calls).
4. **Independent scrutiny of key data** (don't repeat sponsor framing verbatim; flag unverifiable/promotional claims).
5. Ongoing-engagement rules: every number traceable to a source or marked as an estimate; reconcile new documents against the file and flag contradictions/corrections explicitly, including admitting when wrong; no inflating bull or bear case artificially.

This is a **multi-session working file** — the user has already sent one primary-source document (the actual Q1 2026 10-Q) and corrected several placeholder assumptions as a result. More documents (Q2 2026 10-Q, SEABREEZE STAT data, the FY2025 10-K, a proxy/DEF 14A, a Schedule 13D/G on Panacea Venture) are expected in follow-up turns and should be reconciled the same way.

---

## 2. Current deliverables (as of this handoff)

- `/mnt/user-data/outputs/CNTB_Valuation_Model.xlsx` — the workbook. **712 formulas, 0 errors** on last recalc.
- `/mnt/user-data/outputs/CNTB_Investment_Thesis_Memo.docx` — the memo. Rebuilt and verified consistent with the workbook's Base-case numbers as of this handoff.
- All build source is Python (`openpyxl`) for the workbook and Node.js (`docx` npm package) for the memo — see §6 for the full toolchain and how to regenerate/extend either file.

### Headline current outputs (Base case, scenario selector = "Base")
| Metric | Value |
|---|---|
| rNPV per share (Base) | **$5.07** |
| rNPV per share (Low / High) | **$0.74 / $39.34** |
| DCF ceiling case ("if it succeeds," unrisked) | **$16.73** |
| Relative valuation (EV/risk-adj peak sales vs. Upstream Bio) | **$1.08** |
| Blended price target (60% rNPV / 40% RelVal, DCF excluded from blend) | **$3.48** |
| Current share price (as of ~1-Jul-2026 snapshot) | **$2.40** |
| Street consensus (median of 7 analysts, Jul-2026) | **$7.00** |

**⚠ Pending recalc (22-Jul-2026):** the Loss-of-exclusivity assumption was revised via `patch7.py`
(2038/2040/2042 → 2037/2040/2045, see §4 row 11 and §8 correction #9). This environment has no
Excel/LibreOffice to recalculate formulas, so every figure in the table above is still the *prior*
assumption's output. Open the workbook in Excel once and refresh this table (and
`portfolio.json`'s bear/mid/bull) from the real recalculated numbers before relying on them.

---

## 3. Company facts — sourced, as of this handoff

### Identity & history
- **Connect Biopharma Holdings Limited**, incorporated **November 2015** in the Cayman Islands (confirmed via the actual Q1 2026 10-Q — earlier secondary sources were ambiguous between 2012/2015).
- Built around operations in China (Suzhou/Taicang) and a "T-cell-driven" immunology discovery platform.
- IPO'd on Nasdaq **March 2021** at $18.49/share; gross proceeds **$219.9mm** before underwriting discounts (10-Q).
- Original three-candidate pipeline at IPO: CBP-201 (now **rademikibart**) for AD/asthma/CRSwNP; CBP-307/icanbelimod (S1P1 modulator, ulcerative colitis); CBP-174 (H3 antagonist, pruritus).
- **Nov 2021**: global Ph2 in moderate-to-severe atopic dermatitis met primary endpoint; guided toward global Ph3 starting mid-2022.
- **Dec 2022**: company did NOT start the global Ph3 AD program as planned, citing "challenging funding environment," pending a partner.
- **2023**: stock fell ~80% amid disappointing results tied to both lead candidates; CBP-307 shelved entirely, narrowing the pipeline to rademikibart alone.
- **Nuance**: the stalled *global* AD strategy is distinct from China-specific data — the China-only pivotal AD trial (run for Simcere) met all its endpoints around the same period. The 2022-23 reset reads as a funding/strategic retreat from a global AD launch, not a failure of the molecule's biology — but still a real, value-destroying, materialized risk event.
- **2024**: new US-centric leadership (CEO **Barry Quart, Pharm.D.**) refocused the company on rademikibart in **acute exacerbations** of asthma/COPD specifically — a narrower, differentiated, no-approved-biologic-competitor niche — leaving AD to Simcere's China-licensed program.
- Read this as a genuine, twice-repeated pattern: rademikibart has now been re-pointed at three indication families (AD → chronic asthma → acute exacerbations). Could be smart optionality-seeking with a validated molecule and new discipline, or a search for where the data will finally read out clean. Both readings are consistent with the facts.

### Pipeline / catalysts
- **SEABREEZE STAT Asthma** (Ph2, n≈160, acute exacerbations, EOS≥300): enrollment complete 17-Jun-2026. Primary endpoint = 28-day treatment-failure rate; key secondary = post-BD FEV1 at Week 1. **Topline expected early September 2026.**
- **SEABREEZE STAT COPD** (Ph2, n≈160): enrollment substantially complete ~Jun 2026. Topline "shortly after" Asthma — no firm date; **modeled as Q4 2026 (our estimate)**.
- DMC completed a pre-specified interim efficacy review of both studies 23-Apr-2026: no safety concerns, no change to sample size — a positive procedural signal, not a read on efficacy magnitude.
- Supportive prior data: Ph1 IV rademikibart (rapid FEV1 improvement within 24h, sustained ~weeks) and an earlier global Ph2 in chronic moderate-to-severe asthma (published AJRCCM) — topline characterizations only; **we have not pulled the full paper's effect sizes/CIs** (open item, see §7).
- No biologic is approved anywhere today for acute exacerbations of asthma or COPD — genuine white space, but also means the specific endpoint (28-day treatment failure after a single dose) has no precedent for this mechanism.
- COPD carries extra regulatory risk beyond the shared clinical risk: Dupixent's own COPD approval came years after its asthma/AD approvals, in a narrower population — FDA appears more cautious in COPD generally.

### China / Simcere license — full deal economics (confirmed via actual 10-Q, Note 8)
- Agreement signed **21-Nov-2023**. Simcere Pharmaceutical Co., Ltd. holds exclusive Greater China rights (mainland China, HK, Macau, Taiwan), all indications.
- Upfront fee: **~$21mm** (2023).
- Total milestones originally: **up to $123mm**.
- Received in 2024: **~$5mm** (development milestone achieved).
- **LAPSED in 2025: ~$8mm** — a time-based milestone Simcere missed the deadline for. **This is a real, already-materialized partner-execution-risk event**, not a hypothetical — use it as the concrete evidence for "risk already materialized once" in the risk-disclosure sections.
- Remaining eligible milestones as of 31-Mar-2026: **~$110mm**.
- Royalties: tiered, up to **low double-digit %** of Greater China net sales, for **~12 years post-commercialization**.
- China Ph3 AD trial reportedly met all endpoints; NDA submitted to China's regulator. We have NOT independently verified the underlying data package or the "near-maximal responses in ~90% of patients through 52 weeks" characterization — it's a company/partner press-release framing (open item, see §7).
- Q1 2026 revenue recognized: **$169,000** (10-Q exact figure) — amortization of the 2023 upfront fee, not new cash. No revenue in Q1 2025.

### Financials — exact figures from the actual Q1 2026 10-Q (`cntb-20260331.html`, filed **12-May-2026**, period ended 31-Mar-2026)
- **Cash and cash equivalents: $46.034mm.** **Short-term investments: $5.997mm** (separate line). **Total liquidity: $52.03mm.**
  - ⚠️ Correction note: secondary sources (press summaries) quote "$46.0mm" as if it were the *total* — it is actually just the cash-and-equivalents line. Total liquidity is ~$6mm higher than that headline figure.
- **Total debt: $0.** No covenants, no maturity wall. Only operating leases (San Diego HQ, 6,942 sq ft, expires Jan-2028; Taicang China lab/office, 25,476 sq ft, expires Apr-2027; total lease obligations ~$0.7mm — immaterial).
- No material litigation (confirmed, Item 1 Legal Proceedings).
- **Shares issued and outstanding: 62,711,690** at 31-Mar-2026 (vs. 56,442,308 at 31-Dec-2025).
- **Weighted-average shares (basic & diluted), Q1 2026: 56,545,000** (vs. 55,352,000 Q1 2025).
- **Stock options outstanding: 15,167,637** at 31-Mar-2026, **weighted-average exercise price $2.08** — meaningfully in-the-money at the current ~$2.40 share price (real near-term dilution, not a distant hypothetical). Activity in the quarter: 263,800 granted @ $2.50; 131,505 exercised @ $1.04; 155,233 cancelled @ $8.73 (old, high-strike IPO-era grants rolling off).
- **Employee stock purchase plan rights outstanding: 600,000.**
- **⇒ Fully-diluted share count used in the model: ~78.48mm** (62.7117mm basic + 15.1676mm options + 0.60mm ESPP, simple gross if-exercised convention — not treasury-method-reduced). This REPLACES an earlier placeholder of 66.5mm that was a pure guess before this 10-Q was available.
- Q1 2026 income statement: License/collaboration revenue $169K; R&D expense **$15.030mm** (more than 2x Q1 2025's $6.633mm, reflecting SEABREEZE enrollment ramp); G&A **$4.746mm**; total opex $19.776mm; loss from operations $(19.607)mm; other income net $257K; net loss before tax $(19.350)mm; income tax expense $48K; **net loss $(19.398)mm ($0.34/share)** vs. $(10.272)mm ($0.19/share) Q1 2025... exact prior-year comparison label appears as $10,272K net loss (10.27mm) in the primary source, note our memo previously said "$10.3mm" which is consistent rounding.
- Net cash used in operating activities Q1 2026: **$16.0mm** (vs. $10.0mm Q1 2025).
- Net cash provided by financing activities Q1 2026: **$17.6mm** (vs. $2,000 Q1 2025) — driven by the Private Placement.
- Full valuation allowance maintained against all deferred tax assets (confirms conservative 0%-near-term-cash-tax modeling assumption; exact NOL carryforward $ balance NOT disclosed in the 10-Q — would need the FY2025 10-K).
- Single reportable operating segment: "treatment of respiratory diseases." No program-level (Asthma vs. COPD) cost split disclosed.

### March 2026 Private Placement — full detail, including the related-party finding
- Securities purchase agreement, March 2026: **6,130,000 ordinary shares at $3.25/share**. Closed **31-Mar-2026**.
- Gross proceeds **$20.2mm**; estimated net proceeds **$18.6mm** (net of ~$1.6mm placement agent fees/other offering expenses).
- **⚠️ Important correction/finding**: the private placement was **led by Panacea Venture**, described in the 10-Q itself as **"the Company's largest current investor,"** which purchased **$4.0mm** of the $20.2mm round. **Board member James Huang is the sole owner of Panacea Innovation Limited, which is the sole owner of Panacea Venture.** This means the company's own largest shareholder — controlled by a sitting director — anchored its own financing round. This is a **related-party transaction**, not generic PIPE participation, and should not be conflated with ordinary insider buying.
  - Read as a genuinely *stronger* capital-commitment signal than routine PIPE participation (the largest holder underwriting the raise) — but also raises a governance question worth asking directly: were the $3.25/share terms arm's-length, and how was the related-party transaction approved by the board?
  - We do **not** have Panacea Venture's *total* stake size from the 10-Q alone (only this incremental $4.0mm purchase) — **pulling a Schedule 13D/G on Panacea Venture is the single highest-value open research item** (see §7).
- Separately, Director **James Huang personally bought 150,000 shares in the open market on 29-May-2026 at $2.48 (~$372,000)** — a distinct, smaller, genuinely open-market signal, should not be conflated with the PIPE-anchoring role above.

### Ownership — RESOLVED 22-Jul-2026, via Panacea's own Schedule 13D/A filings and the FY2024/FY2025 10-Ks' Item 12 tables (was previously 13D/G-and-13F aggregation, unverified against the 10-Q)

Two distinct, coexisting layers — don't conflate them:

**Concentrated >5% holders** (FY2025 10-K, Item 12, as of 26-Mar-2026, based on 56,521,282 shares
outstanding): Panacea Venture Healthcare Fund II, L.P. (James Huang, director) 12,000,000/21.2%,
BioFortune, Inc. (Dr. Wubin Pan, former President/Chairman) 5,987,431/10.6%, Zheng Wei, Ph.D.
(co-founder/director) 5,015,960/8.9%, Ikarian Capital, LLC 4,933,086/8.7%, Qiming Venture Partners
entities 4,840,898/8.6%, Shanghai Minhui Enterprise Management Consulting Partnership 4,789,758/8.5%,
Advantech Capital II Connect Partnership 4,762,185/8.4%, Lily Asia Ventures entities 3,336,907/5.9%.
These seven holders alone control ~72% of the company. All current officers/directors as a group:
15,944,870/26.6%.

**Panacea/Huang's actual trajectory** (previously the single highest-value open item — now fully
resolved): ~13,760,000 shares (24.9%) as of Mar-2025 (FY2024 10-K) → **12,000,000 shares (21.2%)**
per a Schedule 13D/A filed **26-Nov-2025** (Amendment No. 4 — a real ~1.76mm-share reduction
sometime in 2025, not disclosed anywhere in this file until now) → **13,160,000 shares (21.2%)**
per the Schedule 13D/A filed **1-Apr-2026** (Amendment No. 5), which discloses the exact March 2026
purchase: 1,160,000 shares at **$3.45/share** (not the $3.25/share outside investors paid — the 8-K
itself specifies the $3.25 price applies only to purchasers not affiliated with an officer/director/
employee/consultant; Panacea/Huang paid a premium, consistent with Nasdaq related-party pricing
rules, not a discount). James Huang personally agreed to a 45-day post-closing lock-up (Item 4 of
the 13D/A). Net picture: Panacea trimmed meaningfully in 2025, then used the 2026 raise to rebuild
to roughly its *prior* percentage — a materially different, more nuanced story than "the largest
investor keeps buying," which is what a plain read of the Q1 2026 10-Q's PIPE disclosure alone
would suggest.

**Board change**: James A. Schoeneck is now a director (0 shares as of Mar-2026), replacing the
Qiming-designated Kan Chen, Ph.D. (rotated off sometime between the FY2024 and FY2025 10-Ks).

**Insider alignment, newly visible**: CEO Barry Quart's beneficial ownership grew from 131,887
shares (Mar-2025) to 1,700,473 shares/2.9% (Mar-2026); President David Szekeres's grew from 160,867
to 1,234,840 shares/2.1% — both driven by their 2024 hire-date option grants vesting. Real, growing
skin in the game from the two executives running day-to-day execution into the September data,
not previously reflected in this file.

Separately, a broader 13F aggregation (not re-verified this pass) shows ~59% institutional
ownership across 28 filers holding smaller, sub-5% stakes each: Perceptive Advisors (dedicated
healthcare/biotech long-only — most thesis-relevant), Boothbay Fund Management, BML Capital
Management, Renaissance Technologies, Citadel Advisors, Jane Street Group (last three read as
quant/market-making flow, not directional conviction). Ikarian Capital appears in *both* views,
since its stake is large enough to cross the 5% Schedule 13G threshold (confirmed via a 13G filed
17-Feb-2026) — this cross-validates rather than contradicts the earlier 13F-based note.

Governance nuance, still relevant: CNTB is a **foreign private issuer** per the 10-Q, but has been
a **voluntary SEC domestic filer** (10-K/10-Q/8-K instead of 20-F/6-K) starting with the FY2024
10-K, part of the 2024 U.S.-centric transition. The Holding Foreign Insiders Accountable Act
(HFIAA, signed 18-Dec-2025) newly requires CNTB's officers and directors (not principal
shareholders) to comply with Section 16(a) insider-reporting starting 18-Mar-2026; all remain
exempt from Section 16(b) short-swing profit recovery.

### Analyst coverage (as of Jul-2026 — ratings/targets not re-verified since original research; conflict of interest now CONFIRMED, not unconfirmed)
| Firm | Target | Rating | Date |
|---|---|---|---|
| H.C. Wainwright | $7 | Buy | — |
| BTIG | $10 | Buy | — |
| Cantor Fitzgerald | $4 | Overweight | — |
| Canaccord Genuity | $6 | Buy | — |
| Piper Sandler | $7 | Overweight | — |
| Oppenheimer | $8 | Outperform | initiated 9-Jul-2026 |

Median ~$7.00, aggregate "Strong Buy." Nearly all coverage initiated in the ~10 weeks before the September 2026 data — classic pre-catalyst sell-side positioning; none of these firms has a track record through an actual CNTB data readout. **Placement-agent identity is now confirmed**: the 8-K announcing the March 2026 private placement (filed 30-Mar-2026) names **Leerink Partners LLC and Cantor Fitzgerald & Co.** as joint placement agents. Cantor Fitzgerald is one of the six covering analysts above with a live $4 Overweight rating — a direct, confirmed conflict of interest between arranging a related-party-anchored raise and publishing ongoing sell-side coverage on the same stock. This doesn't necessarily make the rating wrong, but it should be weighted accordingly.

### Relative valuation peers used
- **Upstream Bio (UPB)** — INCLUDED, primary multiple source. Single-asset (verekitug, TSLP-receptor antagonist), Ph2/3, severe asthma + CRSwNP + COPD. EV ~$190mm (Seeking Alpha, 30-May-2026, "under $200M" following a ~47% post-Phase-2-miss selloff — VERIFY against a current cash balance). Risk-adjusted peak WW sales estimate $2.7bn by 2035 (Mizuho, 17-Dec-2025). Implied EV/risk-adj-peak-sales multiple ≈ **0.070x**.
- **Aiolos Bio / GSK (Jan-2024)** — precedent transaction, NOT used in the multiple. $1.0bn upfront + up to $400mm milestones, Ph2-ready single TSLP mAb, asthma-only. Reference/sanity-check only.
- **Apogee Therapeutics / AbbVie (2026)** — EXCLUDED from the multiple. ~$10.9bn all-cash, multi-asset IL-4/IL-13/T2 platform, much larger/more advanced (has an $800mm synthetic royalty + $500mm debt facility, cash into 2029). Included only as evidence of continued large-pharma appetite in the space.

---

## 4. Model assumptions — Low / Base / High (Cover & Assumptions tab, rows 42–57 in the workbook)

All of these are editable input cells (blue fill) with a "Selected" column (light-blue fill) driven by the scenario-selector cell (`Cover & Assumptions!C37`, dropdown Low/Base/High).

| # | Assumption | Low | Base | High | Basis |
|---|---|---|---|---|---|
| 1 | PoS: SEABREEZE Asthma Ph2 positive topline | 40% | 60% | 75% | Judgment — mechanism de-risked (validated IL-4Rα target) but novel endpoint |
| 2 | PoS: Ph2(+)→US Approval, Asthma | 35% | 50% | 65% | Typical Ph2→approval conversion for validated-MOA respiratory biologics |
| 3 | PoS: US Approval, COPD (own binary) | 18% | 28% | 40% | Revised DOWN after confirming the $8mm lapsed China milestone as a partner-execution-risk precedent; COPD approval history for this MOA-class is slower/narrower (Dupixent precedent) |
| 4 | US peak eligible patients — Asthma (000s/yr) | 350 | 450 | 550 | Bottom-up off company's >1mm US ED-visit disclosure, ESTIMATE of EOS/FeNO-eligible subset |
| 5 | US peak eligible patients — COPD (000s/yr) | 400 | 550 | 700 | Bottom-up off CDC-cited ~1.5mm annual COPD-exacerbation ED/hospitalization visits |
| 6 | Net price per treated patient per year ($) | 6,000 | 9,000 | 13,000 | UNKNOWN/ESTIMATE — no company pricing guidance exists; single largest value driver |
| 7 | Peak penetration of eligible population | 12% | 20% | 30% | ESTIMATE — no approved biologic competitor, but a wholly new ED/urgent-care commercial model |
| 8 | US launch year — Asthma | 2032 | 2031 | 2030 | ESTIMATE from data timing |
| 9 | US launch year — COPD | 2033 | 2032 | 2031 | ESTIMATE, ~1yr behind Asthma |
| 10 | Years launch → peak sales | 6 | 5 | 4 | ESTIMATE |
| 11 | Loss-of-exclusivity year | 2037 | 2040 | 2045 | SOURCED 22-Jul-2026 via `build/patch7.py` (was an unverified placeholder). Per FY2024/FY2025 10-Ks, Item 1 IP: composition-of-matter patent family expires 2037 (unextended, controls Low); formulation family expires 2040 (unextended, unchanged Base); Hatch-Waxman PTE could add up to 5yr capped at 14yr-from-approval (High). **NOT YET RECALCULATED** — this environment has no Excel/LibreOffice; open the workbook once to refresh cached rNPV/DCF values before trusting any downstream number in §2's headline table. |
| 12 | COGS, % of net sales | 18% | 15% | 12% | ESTIMATE, typical mAb manufacturing cost |
| 13 | Peak SG&A, % of net sales | 35% | 28% | 22% | ESTIMATE, niche high-touch ED/hospital commercial model |
| 14 | Discount rate / WACC | 14% | 12.5% | 11% | ESTIMATE, standard clinical-biotech range |
| 15 | China milestone realization, % of $110mm remaining | 15% | 30% | 45% | Revised DOWN after confirming the $8mm 2025 lapse — real evidence milestones aren't automatic |
| 16 | China royalty NPV, risk-adjusted, standalone ($mm) | 15 | 30 | 50 | Rough placeholder — no access to Simcere's internal China sales forecast |

**COPD PoS is intentionally set below Asthma's combined PoS** (0.28 vs. 0.60×0.50=0.30) — this was a deliberate recalibration during build to keep the model internally consistent with the stated narrative that COPD is the riskier, slower program (see git-style history in §6 patch log if picking this apart further).

---

## 5. Workbook structure — tab-by-tab, with row references

The workbook has 6 tabs, built via Python/openpyxl, with a full row-index map preserved in `row_map.json` (included in the handoff bundle) so any future edits can target exact cells without re-deriving row numbers by hand.

1. **Cover & Assumptions** — company snapshot (rows 6–17: price/shares/mkt cap/cash/debt/EV), program status & catalysts (rows 19–36ish), scenario selector (row 37), master assumptions table (rows 42–57, columns B=name, C=Low, D=Base, E=High, F=Selected formula, G=source note).
2. **Balance Sheet & Cash Runway** — cash position (rows 7–9), debt (row 11), recent financing & dilution (rows 15–18), share count basic vs. options (rows 22–24), burn rate & runway (rows 28–36), financing-need flag (row 40).
3. **rNPV Model** — 20 year-columns (2026–2045, columns C–V), Asthma block then COPD block (each: years-since-launch, penetration curve, treated patients, revenue, COGS, SG&A, lifecycle R&D, pre-launch Ph3 cost, pretax CF, tax rate, after-tax CF, combined PoS, risk-adjusted CF, discount factor, PV), then a summary block (rows ~47–59: sum of program rNPVs + China milestones/royalty + net cash − burn-drag placeholder = total equity value ÷ FD shares = **rNPV per share, row 59**).
4. **DCF Model** — same operating engine, pulled by direct cell reference from the rNPV tab's *unrisked* after-tax cash-flow rows (no PoS multiplication), explicit terminal value (0% growth, applied to the 2045 post-LOE residual, NOT peak-year CF), summary → **DCF per share, row 37**.
5. **Relative Valuation** — peer screen/inclusion-exclusion logic, UPB multiple derivation, CNTB's own risk-adjusted peak sales (pulled live via `MAX()` off the rNPV tab's revenue rows × combined PoS), implied value → **Relative valuation per share, row 32**. Precedent-transaction sanity checks (Aiolos/GSK, Apogee/AbbVie) shown below, not used in the multiple.
6. **Price Target Summary** — blend table (editable weights, currently 60% rNPV / 40% RelVal, DCF explicitly excluded from the blend), sensitivity table (real one-at-a-time flexes, recomputed after the 10-Q correction), closing caveat on why the Low-High spread is wider than the sum of one-at-a-time sensitivities (compounding uncertainty).

**Formatting conventions used throughout:** blue font + light-yellow fill = input/hardcode; black = formula; green = cross-sheet link; italic grey = source note/caveat; red = flagged genuine unknown. Section headers use a dark-navy fill bar. All monetary values in $mm unless stated per-share.

---

## 6. Toolchain — how to regenerate or extend either file

Both deliverables were built with **code, not manual Excel/Word editing** — this matters for handoff because Claude Code can directly modify the Python/JS source rather than trying to parse/patch the binary files.

### Excel workbook
- Library: `openpyxl` (Python).
- `helpers.py` — shared style constants (fonts, fills, number formats) and a `set_cell()` / `section_bar()` helper. Import from here, never re-run as a script (it has no top-level side effects, safe to import).
- `build1.py` → `build8.py` — sequential build scripts, each appending one tab or block. **Must be run in order on a fresh file** to reproduce the workbook from scratch (each script loads the previous script's saved `model.xlsx`, appends, and re-saves). Order: build1 (Cover title block) → build2 (assumptions table) → build3 (Balance Sheet tab) → **patch1.py** (fixes cross-sheet cell references from Cover to Balance Sheet, needed because exact row numbers aren't known until build3 runs) → build4 (rNPV Asthma block) → build5 (rNPV COPD block + summary) → **patch2.py** (recalibrates net-price/penetration/COPD-PoS assumptions to be more conservative — this correction happened before the 10-Q upload) → build6 (DCF tab) → build7 (Relative Valuation tab) → build8 (Price Target Summary tab).
- `row_map.json` — persisted dictionary of every important row number (see §5), read/written by every build/patch script via `json.load`/`json.dump`. **Always load this before writing new patches** — it's the source of truth for "what row is X on."
- `patch3.py` through `patch6.py` — the corrections applied after the user uploaded the actual Q1 2026 10-Q (`cntb-20260331.html`): patch3 = cash/share-count corrections; patch4 = precision updates + Panacea Venture/James Huang reframing on the Balance Sheet tab; patch5 = Simcere deal economics detail + China-milestone-assumption revision; patch6 = refreshed the Price Target Summary sensitivity table with recomputed deltas after the share-count correction changed the base case.
- `patch7.py` (22-Jul-2026) — revised the Loss-of-exclusivity assumption (row 52) from an unverified placeholder (2038/2040/2042) to real patent-expiry data from the FY2024/FY2025 10-Ks (2037/2040/2045). **Not yet recalculated** — this environment (Claude Code on Windows, no Office/LibreOffice installed) can write input cells via openpyxl but cannot recalculate formulas, so the workbook's cached rNPV/DCF outputs are stale until it's opened in Excel at least once. Don't read Low/Base/High figures out of the cached values until that happens.
- **Recalculation**: the original toolchain used `/mnt/skills/public/xlsx/scripts/recalc.py <file> <timeout>` (a Claude.ai-sandbox-specific LibreOffice-headless recalc-and-error-check) — that path doesn't exist in a plain Claude Code/Windows environment. If no equivalent is available, recalculating means opening the file in Excel (or installing LibreOffice) at least once after any input-cell edit; don't skip this and don't hand-compute a substitute number.
- **Extending further**: to change an assumption going forward, just open `model.xlsx` with openpyxl, write to the relevant Low/Base/High cell in `Cover & Assumptions` (rows 42–57), save, and recalc — all downstream tabs update automatically via formula, no need to touch rNPV/DCF/RelVal/Price-Target tabs directly. To add a wholly new line item (e.g., a real patent-expiry-driven LOE schedule once the 10-K is read), follow the existing pattern: add a new assumption row with Low/Base/High/Selected/Source-note columns, then reference `'Cover & Assumptions'!$F$<row>` from wherever it's needed.
- **Known fragile spot**: in the rNPV Model tab, columns C through V (20 columns) are all *year-data* columns — do NOT write source/note text into any column in that range assuming it's a "notes" column the way other tabs use column G. Notes in that tab go in column X instead (this bit us once during the build — see the `#VALUE!` errors caused by overwriting a formula cell with note text, fixed by moving notes to column X).

### Word memo
- Library: `docx` (npm package), via two Node scripts.
- `memo_partA.js` — helper functions (`h1`, `h2`, `p`, `bullet`, `note`, `cell`, numbering config) plus `titleBlock` and `section1to7` (Executive Summary through Where Outside Expertise Would Help) as exported arrays of `Paragraph` objects. Requires the file to be **valid UTF-8 but stores literal escaped-unicode sequences like `\u201c`/`\u201d` for smart quotes as literal 6-character text in the source** — if patching with Python string-replace, match against those literal escape sequences, not actual curly-quote characters, or the replace will silently no-op (this happened during the correction pass — several `.replace()` calls failed silently until this was diagnosed).
- `memo_partB.js` — imports from `memo_partA.js`, builds the valuation table (`valTable()`), `section8to10` (Valuation Summary through Independent Scrutiny of Key Data), assembles the full `Document`, and calls `Packer.toBuffer().then(...)` to write `CNTB_Investment_Thesis_Memo.docx`. **Run this file (`node memo_partB.js`) to rebuild the docx after any edit to either script.**
- To verify formatting after rebuilding: convert to PDF via `/mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf <file>.docx`, then `pdftoppm -jpeg -r 90 <file>.pdf page` and view the resulting `page-N.jpg` files, or extract plain text via `pandoc -t plain <file>.docx` and grep for key figures to confirm consistency with the workbook.
- **Consistency discipline**: every dollar figure and every fact in the memo must match the workbook. When a workbook number changes (as happened after the 10-Q), grep both `memo_partA.js` and `memo_partB.js` for the old figures before considering the correction complete — this was missed once (Section 8's valuation table and Section 9's sell-side reconciliation still had stale numbers after the first correction pass) and required a follow-up fix.

---

## 7. Open items — ranked by value, for the next research pass

**Resolved 22-Jul-2026** (kept here for the audit trail, not because they're still open): the
Panacea Venture 13D/G pull, the FY2025 10-K's IP/patents and NOL sections, and the private
placement's placement-agent identity — see §3 and §8 for what was found and where.

1. **Recalculate the valuation model** against the revised loss-of-exclusivity assumption
   (patch7.py, 2037/2040/2045) — this environment has no Excel/LibreOffice, so the model's cached
   rNPV/DCF outputs (and this doc's §2 headline table) still reflect the *prior* assumption. Open
   the workbook in Excel once, then update §2 and `portfolio.json`'s bear/mid/bull from the real
   recalculated numbers — don't hand-type a guess.
2. **Q2 2026 10-Q** (due ~12-Aug-2026) — tests whether the tighter mechanical-burn-extrapolation
   runway estimate or management's more conservative-than-thought "≥12 months from 12-May-2026
   filing" (~May-2027) framing is closer to right.
3. **Full AJRCCM paper** (not just the press-release characterization) for the prior
   chronic-asthma Ph2 — actual effect sizes/CIs would sharpen how much read-through to give the
   SEABREEZE endpoint. (Note: this is separate from RADIANT-AD and the Phase 1 IV study, both of
   which now have real disclosed numbers — see §3.)
4. **RADIANT-AD verification** — current 52-week China Ph3 AD data (EASI-75/IGA-0/1/EASI-90) is a
   company 8-K exhibit presented at the March 2026 AAD meeting, not yet a peer-reviewed publication
   or NMPA-level data package. Revisit once either lands.
5. **Current DEF 14A / proxy** — exec comp detail specifically (board composition itself is now
   resolved via the FY2024/FY2025 10-Ks, see §3).
6. **SEABREEZE STAT topline data itself** (expected early Sept 2026 for Asthma) — the big one; when
   it lands, the entire PoS-driven structure of this model should be revisited, not just patched (a
   positive readout should probably collapse the Low/Base/High PoS-for-Ph2 assumption to ~100% and
   shift focus entirely to pricing/penetration/Ph3-execution risk).

---

## 8. Corrections log (for auditability — "own your mistakes" per the engagement rules)

| # | What was wrong | What corrected it | Source |
|---|---|---|---|
| 1 | Fully-diluted shares assumed ~66.5mm (placeholder guess) | Corrected to ~78.48mm (62.71mm basic + 15.17mm options @ $2.08 strike + 0.6mm ESPP) | Actual Q1 2026 10-Q, Note on Net Loss per Share + Note 10 Equity Incentive Plans |
| 2 | Total liquidity modeled as $46.0mm (cash $30.1mm + ST invest $15.9mm, a guessed split) | Corrected to $52.03mm total ($46.034mm cash + $5.997mm ST investments — the "$46.0mm" figure in press summaries was only ever the cash-and-equivalents line) | Actual 10-Q balance sheet |
| 3 | James Huang's $4.0mm March 2026 purchase framed as generic "PIPE participation" | Reframed as Panacea Venture (Huang-controlled, company's largest shareholder) leading/anchoring the entire private placement — a related-party transaction | Actual 10-Q, Note 9 Shareholders' Equity |
| 4 | China milestone realization probability set at 20/35/50% with no specific negative precedent cited | Revised down to 15/30/45% after confirming a real $8mm milestone lapse in 2025 (Simcere missed a deadline) | Actual 10-Q, Note 8 License and Collaboration Agreement |
| 5 | Company incorporation date ambiguous (2012 vs. 2015 across secondary sources) | Confirmed November 2015, Cayman Islands | Actual 10-Q |
| 6 | Going-concern runway framed only as management's "into 2H 2027" PR language | Added the more conservative formal 10-Q commitment: "at least one year from the 12-May-2026 filing date" (~May-2027 floor) | Actual 10-Q, Note 2 Liquidity and Going Concern |
| 7 | Options/RSU overhang treated as an unknown ~3.5mm placeholder | Resolved to the exact 15,167,637 options (WA strike $2.08) + 600,000 ESPP rights | Actual 10-Q |
| 8 | Panacea Venture's March 2026 $4.0mm purchase framed as "largest current investor" with total stake unknown | Reframed with full trajectory: ~13.76mm sh (24.9%) Mar-2025 → 12.0mm sh (21.2%) per 13D/A filed 26-Nov-2025 (a real ~1.76mm-share reduction, not previously known) → 13.16mm sh (21.2%) per 13D/A filed 1-Apr-2026, buying 1,160,000 sh at $3.45/share (a premium to the $3.25 outside price, not a discount) | Panacea's own Schedule 13D/A filings (SEC EDGAR, accessions 0001193125-25-300640 and 0001193125-26-138013) |
| 9 | Loss-of-exclusivity Low/Base/High (2038/2040/2042) was an explicitly-flagged unverified placeholder | Sourced from actual patent data: composition-of-matter family expires 2037 (unextended), formulation family expires 2040 (unextended), Hatch-Waxman PTE up to 5yr capped at 14yr-from-approval. Revised to 2037/2040/2045 via patch7.py — **not yet recalculated through the model** | FY2024 10-K and FY2025 10-K, Item 1 Intellectual Property (both filings, consistent figures) |
| 10 | NOL carryforwards noted as "not disclosed in the 10-Q" | Fully disclosed in the 10-K: China NOLs $193.3mm (begin expiring 2026), Hong Kong $225.8mm (indefinite), Australia $5.0mm (indefinite), US federal R&D credits $5.4mm (begin expiring 2042). Full valuation allowance zeroes net DTA regardless — confirms the existing 0%-near-term-tax assumption was already right | FY2025 10-K, Note 11 Income Taxes |
| 11 | Sell-side placement-agent conflict of interest: "could not confirm from public sources" | Confirmed: Leerink Partners LLC and Cantor Fitzgerald & Co. were joint placement agents on the March 2026 raise. Cantor Fitzgerald carries a live $4 Overweight rating on CNTB — a direct, confirmed conflict | 8-K filed 30-Mar-2026, Item 1.01 |
| 12 | "~90% near-maximal responses" China AD characterization flagged as an unverified press-release claim | Traced to a real, previously-unnamed 259-patient Phase 3 trial, RADIANT-AD: 52-week data shows EASI-75 96.6%, IGA 0/1 87.1%, EASI-90 85.3% — the "~90%" characterization is a fair average of real figures, not marketing inflation, though still a company topline, not peer-reviewed | 8-K exhibit 99.2, filed 30-Mar-2026 (RADIANT-AD topline press release) |
| 13 | Phase 1 IV study effect size known only as a vague "rapid and significant improvement" characterization | Quantified: mean FEV1 improvement ~200-400 mL maintained through Day 29 in asthma and COPD patients (100-400 mL increases as early as 15 min post-dose); small study (12 asthma/10 COPD, 4:1 randomized) - supportive mechanistic signal, not powered for efficacy | 8-K exhibit 99.1, filed 30-Mar-2026 (CBP-201-105 topline press release) |
| 14 | Board composition: "did not find a detailed current board composition" | Resolved via both 10-Ks: Kleanthis Xanthopoulos (Chair), James Huang, Zheng Wei, Wubin Pan, Barry Quart remain; Kan Chen (Qiming designee) rotated off, replaced by new director James A. Schoeneck | FY2024 10-K and FY2025 10-K, Item 12 |
| 15 | CEO/President personal share ownership treated as static (131,887 / 160,867 shares) | Both grew sharply as 2024 hire-date options vested: Quart to 1,700,473 sh (2.9%), Szekeres to 1,234,840 sh (2.1%), as of Mar-2026 | FY2025 10-K, Item 12 |

All corrections through #7 were reconciled against the workbook and memo and re-verified error-free
at the time. Corrections #8–15 (22-Jul-2026) update `portfolio.json`, `README.md`, and the memo
build scripts (`memo_partA.js`/`memo_partB.js`, rebuilt via `node memo_partB.js`); #9 (loss of
exclusivity) additionally touches the workbook via `patch7.py` but has **not** been recalculated
through to the rNPV/DCF outputs — see §7, item 1.

All six downstream valuation outputs (rNPV, DCF, Relative Valuation, Price Target blend) were recomputed and the workbook re-verified error-free after each correction; the memo was re-grepped and rebuilt to match.

---

## 9. Files included in this handoff bundle

- `CNTB_Valuation_Model.xlsx` — current workbook (0 formula errors, matches all figures in this document)
- `CNTB_Investment_Thesis_Memo.docx` — current memo (consistency-checked against the workbook)
- `row_map.json` — row-index map, required for any further programmatic edits
- `helpers.py`, `build1.py`...`build8.py`, `patch1.py`...`patch6.py` — full Python/openpyxl build history for the workbook
- `memo_partA.js`, `memo_partB.js` — full Node/docx build history for the memo
- `cntb-20260331.html` — the actual Q1 2026 10-Q as originally uploaded by the user (primary source for §3's financial facts)
- `Biotech_DD_Prompt_Template.md` — the user's original reusable prompt template that this whole engagement follows

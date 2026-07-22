# Biotech Research Repo — Context for Claude Code

This repo is the output of an extended research/valuation project done in a Claude.ai chat, being
migrated to Claude Code so files can be committed and PRs opened directly. Read this whole file
before doing anything — it's the brief a fresh session needs since none of the original chat
history carries over.

## What this repo is for

A personal research operation for tracking biotech stock theses. Each company gets a folder with
a full valuation model, a written thesis memo, and its source materials. A `portfolio.json` at the
repo root acts as an index that a dashboard (`dashboard/portfolio_monitor.html`) reads to show all
positions at a glance.

## Repo layout

```
biotech-research/
├── portfolio.json                     <- index the dashboard reads (one row per company)
├── dashboard/
│   ├── portfolio_monitor.html         <- primary dashboard: fetches portfolio.json from GitHub's raw URL
│   └── portfolio_ledger.html          <- earlier version, browser-local-storage only, no GitHub fetch (keep as fallback reference)
├── templates/
│   ├── Biotech_DD_Prompt_Template.md   <- reusable prompt for kicking off a new company's full analysis
│   └── News_Update_Prompt_Template.md  <- reusable prompt for reconciling news/filings against an existing thesis
├── scripts/
│   └── refresh_prices.py              <- pulls live quotes into portfolio.json's price fields only (see below)
├── .claude/skills/thesis-update/       <- project skill wrapping News_Update_Prompt_Template.md for Claude Code
└── companies/
    ├── SPRB/
    │   ├── SPRB_Valuation_Model.xlsx  <- rNPV / DCF / relative valuation / balance sheet model
    │   ├── SPRB_Thesis_Memo.docx      <- long thesis, risks, ownership, valuation summary
    │   └── source_materials/          <- original documents that informed the analysis (see below)
    └── CNTB/
        ├── CNTB_Valuation_Model.xlsx      <- rNPV / DCF / relative valuation / balance sheet model
        ├── CNTB_Investment_Thesis_Memo.docx
        ├── CNTB_Project_Context.md        <- full handoff doc: corrections log, exact row map, open items
        ├── build/                         <- openpyxl/docx build scripts that generate the two files above
        └── source_materials/               <- primary-source filings (10-Q etc.)
```

## First tasks for this session — DONE, see below for resolution

1. ~~Initialize git if not already done, commit everything as-is, push to the user's GitHub repo.~~
2. ~~Confirm the repo is set to **public**~~ — confirmed public via `gh`/GitHub API.
3. ~~Open `dashboard/portfolio_monitor.html` in a browser and sanity-check~~ — confirmed working
   live against `raw.githubusercontent.com/kuda4910/biotech-research/main/portfolio.json` in a
   real browser (see fix below).
4. ~~Ask the user whether they want future updates via direct commits or PRs~~ — settled: **direct
   commits to main**, no PR review step.

## SPRB (Spruce Biosciences) — full working context

This is the one fully-built example. Everything below is what a new session needs to keep working
on it correctly, without re-deriving mistakes that were already caught and fixed.

### The company

Spruce Biosciences (NASDAQ: SPRB) is a single-asset, late-stage biopharmaceutical company. Lead
asset: **tralesinidase alfa (TA-ERT)**, an intracerebroventricularly-delivered enzyme replacement
therapy for **MPS IIIB (Sanfilippo Syndrome Type B)**, an ultra-rare fatal pediatric
neurodegenerative disease with no approved treatment. FDA Breakthrough Therapy Designation granted
Oct 2025. BLA filing guided for Q4 2026 (already slipped once from an original Q1 2026 target, due
to FDA CMC-focused meetings).

### Origin story (important — explains the deal economics)

TA-ERT was originally developed by **BioMarin**, out-licensed to **Allievex Corporation** in 2019.
Allievex ran the clinical program for years (the long-term cognition data in the thesis originates
from their work), but could not get FDA alignment on an accelerated-approval biomarker pathway and
folded in 2023. In March 2024 the FDA reversed course and accepted HS-NRE as a valid surrogate
biomarker — too late for Allievex. Spruce acquired the program via an April 2025 Asset Purchase
Agreement, assuming the original BioMarin license terms largely as-is: **up to $122.5mm in
milestones + a "high-single-digit to low-teens" royalty on worldwide net sales (modeled at 10%
base case)**. In the same month as that acquisition, Spruce also disclosed a 55% workforce
reduction, a Nasdaq delisting notice, and a going-concern warning — this company was in genuine
financial distress immediately before the asset that now drives its entire valuation arrived.
1-for-75 reverse stock split effective August 2025.

### The valuation model (SPRB_Valuation_Model.xlsx)

Seven tabs, fully formula-driven, zero recalc errors as of last build:

1. **Cover & Assumptions** — company snapshot + a Low/Base/High assumptions panel with a single
   scenario-selector cell (`B43`) that drives every downstream tab.
2. **Balance Sheet & Runway** — cash, Avenue Capital debt terms, share count build.
3. **rNPV Model** — epidemiology-driven patient forecast (this is the single widest-uncertainty
   input in the whole file — patient population is a first-principles estimate, not KOL-confirmed)
   → revenue → risk-adjusted (78% base-case probability of commercial success) → discounted, NO
   terminal value. **Base-case output: ~$19/share.**
4. **DCF Model** — same engine, unrisked "if it works" case, WITH a terminal value.
   **Base-case output: ~$29/share.**
5. **Relative Valuation** — a properly-derived EV/Peak-Sales multiple from Zevra Therapeutics
   (~3.35x, the closest true economic analog — a recurring-dose ultra-rare therapy, not a one-time
   gene-therapy cure like Ultragenyx's UX111, which was deliberately excluded as a multiple source).
   Applied to SPRB's own projected peak revenue. **Base-case output: ~$165/share, bull case ~$383.**
6. **Price Target Summary** — editable-weight blend of the three methods (~$51 blended) plus a
   log of current sell-side analyst coverage (11 analysts, avg target ~$160-163, range $120-259).
7. **Sources** — citations and known data gaps.

**Known bugs already caught and fixed in this file** (don't reintroduce): a formula once
referenced Zevra's market cap instead of its enterprise value (inflated the multiple to 4.79x
instead of 3.35x); a `note()` helper call once silently overwrote a probability-of-success cell
with an empty string, zeroing out the 2026 pre-launch cost in the rNPV. Both fixed and verified.

**Current share count: confirmed at ~2,752,278 basic / ~2,962,278 fully-diluted**, cross-validated
against two independent sources (Godel Terminal, Morningstar) after an earlier back-and-forth where
a misread TradingView chart briefly (incorrectly) suggested the share count hadn't grown post the
April 2026 $69mm offering. It had — Godel/Morningstar were right, the TradingView read was wrong.

### The thesis memo (SPRB_Thesis_Memo.docx)

9 sections: Executive Summary, Long Thesis (incl. 2.3 the Allievex origin story), Key Risks,
Balance Sheet Summary, Institutional Ownership & Insiders, Where Outside Expertise Would Help,
Valuation Summary, an **Appendix (Section 8) independently verifying a third-party statistical
reconstruction** of the WORLDSymposium cognition data (by "Randall Ching" — see source_materials),
and a Disclaimer.

**The Ching appendix matters**: independently re-derived his Hedges' g / pooled-SD calculations
and they check out exactly. Conclusion: the cognitive effect is real but modest (Hedges' g ~0.13-
0.48, not large), concentrated in a specific 7.5-9.0 year age window, with a plausible narrower-
than-all-comers label risk if that holds up under FDA scrutiny. This sharpens (not invents) the
confirmatory-trial risk already in Section 3.1.

### Known corrections already made (don't re-flag as new issues)

- The tildacerfont/HMNC partnership for Major Depressive Disorder (TAMARIND trial) was
  **discontinued in Q1 2026** after a serious adverse event (liver enzyme elevation) — this is
  a different molecule from TA-ERT (oral small molecule vs. biologic ERT), no safety read-through,
  but it means that program is no longer a source of "unmodeled optionality" as once assumed.
- SPR202 (CAH follow-on asset, licensed from HMNC) carries up to $390mm in potential milestones —
  noted but still not modeled.
- New institutional holder found via 13G: Millennium Management / Israel Englander (~103,000 shares).

### Open questions flagged for the user's own research contacts (highest-value first)

1. Real diagnosed MPS IIIB patient counts by country — the single biggest lever in the model.
2. Whether approval could carry an age/stage label restriction given the Ching findings.
3. Whether the confirmatory trial (5yr, placebo-controlled, n≈14) is actually powered to detect a
   Hedges' g ~0.3-0.4 effect.
4. Exact BioMarin royalty rate (disclosed only as a range) — likely findable in the April 2025
   asset purchase agreement exhibit on EDGAR.
5. PRV (Priority Review Voucher) program reauthorization status — a Congressional/legislative
   question, not a scientific one.
6. Avenue Capital loan's exact tranche-release conditions.

### Context on public commentary (fact-checked, don't re-litigate as new)

Martin Shkreli publicly disclosed a long position in SPRB (Oct 2025, X posts + a YouTube video),
price target $500, later described as bought at $30 and added at $50-60. His stated heuristic
("rare disease sole-treatment companies go to $1.5-4bn market cap") is real and grounded (Amicus
Therapeutics was just acquired by BioMarin at $4.55bn) but every clean example reached that tier by
becoming multi-product or by being acquired — not by single-asset organic re-rating, which is the
harder path and the one he explicitly ruled out ("I don't think Spruce will get bought by
BioMarin") when asked directly. He also misattributed TA-ERT's origin to Ultragenyx in that same
answer — factually wrong (it's Allievex/BioMarin, see above) — worth weighing against how much
credit to give his other public claims.

## CNTB (Connect Biopharma) — full working context

Second fully-built example, onboarded from a separate Claude.ai handoff bundle. Full detail lives
in `companies/CNTB/CNTB_Project_Context.md` (corrections log, exact row map, toolchain notes) —
this section is the condensed version for a fresh session to work from without reading that whole
file first.

### The company

Connect Biopharma Holdings Limited (NASDAQ: CNTB), incorporated Nov 2015, Cayman Islands, IPO'd
Nasdaq March 2021. Lead (now only) asset **rademikibart**, an anti-IL-4Rα monoclonal antibody,
now positioned for **acute exacerbations of asthma and COPD** — a niche with no approved biologic
competitor anywhere today. This is the molecule's **third** indication reposition: originally
atopic dermatitis (global Ph3 never started post-2022 funding retreat, though the China-only AD
trial for partner Simcere did meet its endpoints), then chronic asthma, now acute exacerbations
under 2024-installed US-centric leadership (CEO Barry Quart). Read as either disciplined
optionality-seeking on a validated target, or a search for a clean readout — both fit the facts.

### Pipeline / the binary event this model hinges on

**SEABREEZE STAT Asthma** (Ph2, n≈160, acute exacerbations): enrollment complete 17-Jun-2026,
**topline expected early September 2026**. **SEABREEZE STAT COPD** (same size): topline modeled
as Q4 2026 (our estimate, no firm company date). A clean DMC interim safety review (23-Apr-2026,
no sample-size change) is a procedural positive only, not an efficacy signal. When Asthma topline
lands, the whole PoS-driven model structure should be revisited (not just patched) — a positive
readout should collapse the Ph2 PoS assumption toward ~100% and shift focus to pricing/penetration.

### The valuation model (CNTB_Valuation_Model.xlsx)

Six tabs (Cover & Assumptions, Balance Sheet & Runway, rNPV, DCF, Relative Valuation, Price Target
Summary), 712 formulas, 0 recalc errors as of handoff. Base case:

| Method | Output |
|---|---|
| rNPV (Low / Base / High) | **$0.74 / $5.07 / $39.34** |
| DCF ceiling (unrisked, "if it works") | **$16.73** |
| Relative valuation (EV/risk-adj peak sales vs. Upstream Bio, ~0.070x) | **$1.08** |
| Blended target (60% rNPV / 40% RelVal, DCF excluded) | **$3.48** |
| Current price (~1-Jul-2026 snapshot) | **$2.40** |
| Street consensus (median of 6 analysts) | **$7.00** |

Note the relative-valuation multiple comes in **below** rNPV here (opposite of SPRB, where RelVal
was the bull case) — Upstream Bio's post-Ph2-miss selloff depressed its EV/peak-sales multiple to
~0.070x, so applying it to CNTB's own peak sales yields a low implied value. `portfolio.json` uses
the model's own rNPV Low/Base/High as bear/mid/bull (the most internally-consistent single-method
range), not a cross-method mix.

**Built entirely via code** (unlike SPRB): `companies/CNTB/build/` has the openpyxl scripts
(`build1.py`...`build8.py`, `patch1.py`...`patch6.py`, `helpers.py`, `row_map.json`) and the
docx scripts (`memo_partA.js`, `memo_partB.js`) that generate the two deliverable files. To change
an assumption, edit the relevant Low/Base/High cell in Cover & Assumptions directly in the xlsx (or
add a new patch script) — downstream tabs update via formula. **Fragile spot**: in the rNPV tab,
columns C–V are year-data columns; put notes in column X, not into any C–V cell (this caused
`#VALUE!` errors once during the build).

### Known corrections already made (don't re-flag as new issues)

- Fully-diluted share count: was a placeholder ~66.5mm guess → confirmed **~78.48mm** (62.71mm
  basic + 15.17mm options @ $2.08 WA strike + 0.6mm ESPP) from the actual Q1 2026 10-Q.
- Total liquidity: was modeled as $46.0mm on a guessed cash/ST-investment split → corrected to
  **$52.03mm** ($46.034mm cash + $5.997mm short-term investments) — press summaries quoting
  "$46.0mm" as the total are wrong, that's only the cash line.
- **March 2026 private placement ($20.2mm gross, $3.25/share) was led by Panacea Venture**,
  described in the 10-Q as CNTB's largest current investor. Board member James Huang is the sole
  owner of Panacea Innovation Limited, which owns Panacea Venture — so this is a **related-party
  transaction** (Panacea bought $4.0mm of the round), not generic PIPE participation. Don't
  conflate with Huang's separate, smaller open-market purchase (150,000 shares, $2.48, 29-May-2026).
- China milestone realization probability revised **down** (20/35/50% → 15/30/45%) after
  confirming a real, already-materialized precedent: Simcere missed a milestone deadline and
  **~$8mm lapsed in 2025**. Remaining eligible China milestones: ~$110mm of the original $123mm.
- Company incorporation date confirmed **November 2015** (secondary sources were ambiguous
  2012 vs. 2015).

### Open questions flagged for the user's own research (highest-value first)

1. **Schedule 13D/G on Panacea Venture** — resolves its *total* stake beyond the $4.0mm PIPE
   purchase; likely the single most important ownership fact currently missing.
2. FY2025 10-K — patent/IP section (resolves the 2038/2040/2042 loss-of-exclusivity estimate,
   currently a guess) and the NOL carryforward footnote.
3. Q2 2026 10-Q, due ~12-Aug-2026 — tests the mechanical burn-extrapolation runway estimate
   against management's "≥12 months from the 12-May-2026 filing" framing (~May-2027 floor).
4. Full AJRCCM paper for the prior chronic-asthma Ph2 (only have the press-release
   characterization of effect sizes so far).
5. Private placement Form D / placement-agent identity — the ~$1.6mm placement-agent fee is
   confirmed in the 10-Q, but not who received it (sell-side conflict-of-interest question).
6. Current DEF 14A/proxy — board composition, founder's remaining influence, exec comp.

### Analyst coverage (as of Jul-2026, not re-verified since original research)

H.C. Wainwright $7 Buy, BTIG $10 Buy, Cantor Fitzgerald $4 Overweight, Canaccord Genuity $6 Buy,
Piper Sandler $7 Overweight, Oppenheimer $8 Outperform (initiated 9-Jul-2026). Median ~$7, all but
one initiated in the ~10 weeks before the September data — pre-catalyst positioning, no firm here
has a track record through an actual CNTB readout. Placement-agent identity for the March 2026
raise is unconfirmed against this list (see open questions above).

## Workflow tooling

### Starting a new company (templates/Biotech_DD_Prompt_Template.md)

Paste it into a new chat (or adapt it for a Claude Code session) with the ticker filled in. It
specifies the same model structure, memo structure, sell-side reconciliation, and — importantly —
a standing instruction to flag every assumption's source, verify its own math, and self-correct
visibly rather than bury mistakes.

### Updating an existing thesis (templates/News_Update_Prompt_Template.md + the thesis-update skill)

When new information comes in about SPRB, CNTB, or a future position — a filing, an article, a
screenshot, a claim — don't just discuss it in isolation; reconcile it against the existing
thesis. `templates/News_Update_Prompt_Template.md` is the standalone version of this process
(portable to any chat); `.claude/skills/thesis-update/` wraps the same process as a Claude Code
project skill so it's directly invocable in this repo without pasting the whole template each
time. Either path: loads current state (README section + `portfolio.json` entry +
`*_Project_Context.md` if present) → sorts new facts into confirms/corrections/new-risks/
assumption-changes → surfaces materiality and asks before editing → on confirmation, keeps
`portfolio.json`, the README section, the model, and the memo in sync → logs the correction
plainly → commits straight to `main`. This is exactly the process that produced the CNTB
corrections log (§8 in its Project Context doc), now made repeatable instead of ad hoc.

### Refreshing prices (scripts/refresh_prices.py)

```
python scripts/refresh_prices.py            # updates portfolio.json's price + updated fields
python scripts/refresh_prices.py --dry-run  # preview only, writes nothing
```

Pulls live quotes from Yahoo Finance's public chart endpoint (no API key, just a browser-like
User-Agent header — the endpoint 429s without one). Deliberately narrow: it only ever touches
`price` and `updated`, never thesis text, bear/mid/bull targets, catalysts, or open questions —
those require judgment and belong in the news-update process above, not a script. It doesn't
touch git either; review the diff and commit/push yourself. Not currently scheduled to run
automatically — a background job silently editing tracked data with no visible reasoning trail
would undercut the one thing this whole repo is for (every number traceable to a source);
run it manually, or ask to wire it into this environment's scheduled-task support if you want it
on a cadence.

## Known limitations — resolved

- `portfolio_monitor.html`'s live GitHub fetch was untested at handoff because it was built inside
  a Claude.ai artifact, which provides `window.storage` for persistence but doesn't exist outside
  that sandbox. Fixed in commit `0ecac0f`: `getCachedUrl`/`setCachedUrl`/`getCachedData`/
  `setCachedData` now fall back to `localStorage` when `window.storage` is undefined. Verified
  end-to-end in a real browser against the repo's raw URL — fetch, KPI row, and detail-row expand
  all work.
- Git/GitHub access from Claude Code is confirmed working (credential manager already configured);
  commits push straight to `main` per the direct-commits decision above.
- There's a stray empty file named `git` (0 bytes) at the repo root, untracked — noise from some
  earlier command, not part of the repo structure. Left alone; safe to delete whenever.

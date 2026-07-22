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
│   └── Biotech_DD_Prompt_Template.md  <- reusable prompt for kicking off a new company's full analysis
└── companies/
    └── SPRB/
        ├── SPRB_Valuation_Model.xlsx  <- rNPV / DCF / relative valuation / balance sheet model
        ├── SPRB_Thesis_Memo.docx      <- long thesis, risks, ownership, valuation summary
        └── source_materials/          <- original documents that informed the analysis (see below)
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

## The reusable template (templates/Biotech_DD_Prompt_Template.md)

Use this to kick off a new company folder from scratch — paste it into a new chat (or adapt it for
a Claude Code session) with the ticker filled in. It specifies the same model structure, memo
structure, sell-side reconciliation, and — importantly — a standing instruction to flag every
assumption's source, verify its own math, and self-correct visibly rather than bury mistakes.

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

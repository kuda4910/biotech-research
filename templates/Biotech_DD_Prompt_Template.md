# Biotech Deep-Dive Prompt Template

Fill in the bracketed fields and paste this as your opening message for a new company.

---

I'd like to make a full investment analysis on **[COMPANY NAME] ([TICKER])**, focused on **[LEAD ASSET / DRUG NAME]** for **[INDICATION]**. You will be my assistant analyst. Please treat this as an ongoing working file — I'll come back over multiple sessions with follow-up documents, screenshots, articles, and questions to refine it.

## 1. Valuation model (Excel, fully formula-driven — no hardcoded outputs)

Build one workbook with these tabs:

- **Cover & Assumptions** — company snapshot (ticker, price, market cap, shares outstanding, regulatory status, key dates) plus a master assumptions panel with **Low / Base / High columns and a single scenario-selector cell** that flows through every downstream tab. Flag every assumption's source, and mark genuine unknowns clearly (don't silently guess and present it as fact).
- **Balance Sheet & Cash Runway** — cash, debt (with actual terms: covenants, tranching, maturity), burn rate, runway into a specific future date, recent financings and dilution, share count (basic and fully-diluted, with the actual components of the gap between them).
- **rNPV Model** — bottom-up patient population (epidemiology-driven, not just "analyst consensus"), pricing, revenue build, cost of goods/SG&A, any licensing royalties or milestones owed to a partner, probability-of-success risk-adjustment, discounted to a per-share value. This should be the most conservative, most defensible number in the file.
- **DCF Model** — same operating engine, unrisked "if it succeeds" case, with an explicit terminal value. This is the ceiling case, not the base case — label it that way.
- **Relative Valuation** — find genuine peer companies (similar disease category, similar stage, similar business model) and derive an actual **EV/Peak-Sales multiple** from how the market prices them, then apply that multiple to *this* company's own projected peak revenue. Do not just borrow a peer's total market cap and divide by this company's share count — that's a share-count artifact, not a valuation. Explicitly explain why any excluded peer was excluded (e.g., different revenue-recognition profile, diversified vs. single-asset, no clean way to isolate the relevant EV).
- **Price Target Summary** — blend the methods with editable weights, show upside/downside vs. current price, and include a short sensitivity section on what would move the number most.

After building it, **recalculate the workbook and confirm zero formula errors** before showing it to me. Check your own cell references — if a number looks internally inconsistent (e.g., one method's ranking doesn't make sense against another), trace it down rather than presenting it as-is.

## 2. Written thesis memo (Word doc)

- **Executive summary**
- **The long thesis** — why this could work, and honestly, why the market might be under- or over-pricing it right now
- **Origin story / corporate history** — where did the key asset actually come from? Was it developed in-house, licensed, or acquired out of a distressed or failed company? Any prior safety signals, trial failures, workforce cuts, going-concern warnings, or delistings in this company's own history? This kind of history often explains deal terms and risk that aren't obvious from the current numbers alone.
- **Key risks** — regulatory/clinical, commercial/market, financial/capital-structure, ownership/governance. Where a risk has already materialized once (to this asset, a predecessor company, or a close comparator), say so explicitly rather than treating it as a generic category.
- **Balance sheet & cash runway summary**
- **Institutional ownership & insider activity** — who actually holds this, and is recent insider activity open-market buying (real conviction) or just options/RSU/ESPP mechanics (routine, not a signal)?
- **Where outside expertise would help most** — a ranked list of the specific things I could get a better answer on from real contacts (physicians, KOLs, industry professionals) than from public filings alone.
- **Valuation summary table** comparing all methods plus current price and Street consensus.

## 3. Sell-side & third-party research reconciliation

- Pull current analyst coverage: ratings, price targets, and — where available — their stated methodology or rationale.
- Compare their cluster of targets to each of my own model's outputs, and tell me honestly which of my methods it validates or contradicts, and why.
- Flag conflicts of interest — is any covering analyst's firm also an underwriter on this company's financings?
- If any public figure (analyst, well-known investor, YouTuber, etc.) has made a specific public call on this stock, look into it: what's the actual claim, what's it based on, does it hold up, and is it still current or stale.

## 4. Independent scrutiny of key data

If there's a specific piece of clinical, scientific, or financial data the thesis hinges on (a trial endpoint, an effect size, a biomarker claim), and if you can access the underlying summary statistics, do your own independent check rather than repeating the sponsor's framing verbatim — recompute what you can, and tell me plainly if a claim is well-supported, overstated, or unverifiable from what's public.

## 5. Rules for the whole engagement

- Every number in the model should be traceable to a source or clearly marked as an estimate. Never present a guess as a filed fact.
- When you're not sure about something (share count, a deal term, a multiple), say so and tell me what would resolve it.
- If I show you a screenshot, a chart, a slide deck, or a claim from somewhere, actually reconcile it against the model — update the file if it changes something, and tell me plainly if it contradicts something you had before, including if that means you were wrong.
- Don't inflate the bull case to make the stock look more attractive, and don't manufacture bearishness either — I want the real range of outcomes, including the ones that make me want to walk away.
- If you make a mistake in the model (wrong cell reference, wrong unit, wrong row), own it clearly, fix it, and tell me what changed and why — don't bury the correction.

---

**Company-specific context to fill in before sending:**
- Ticker / ideally an investor-deck link, recent 10-K/10-Q, or an SEC filing index
- Any research you've already gathered (analyst notes, forum posts, papers, etc.) — attach or paste
- Anything you already suspect or have heard secondhand that you want fact-checked

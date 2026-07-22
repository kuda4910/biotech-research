const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  LevelFormat
} = require("docx");
const fs = require("fs");

const NAVY = "1F3864";
const GREY = "595959";
const LIGHTBLUE = "D9E1F2";

function h1(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 } });
}
function h2(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } });
}
function p(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, italics: opts.italics || false, bold: opts.bold || false, size: opts.size || 21 })],
    spacing: { after: 160, line: 276 },
    alignment: opts.align || AlignmentType.LEFT,
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    children: [new TextRun({ text, size: 21 })],
    numbering: { reference: "bullets", level },
    spacing: { after: 90, line: 276 },
  });
}
function note(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 19, color: GREY })],
    spacing: { after: 200, line: 260 },
  });
}
function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 20, type: WidthType.PERCENTAGE },
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: opts.shade } : undefined,
    children: [new Paragraph({
      children: [new TextRun({ text, bold: opts.bold || false, size: 20, color: opts.color || "000000" })],
      alignment: opts.align || AlignmentType.LEFT,
    })],
    verticalAlign: "center",
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
  });
}

const numbering = {
  config: [{
    reference: "bullets",
    levels: [
      { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 260 } } } },
    ],
  }],
};

const titleBlock = [
  new Paragraph({
    children: [new TextRun({ text: "Connect Biopharma Holdings Limited (NASDAQ: CNTB)", bold: true, size: 40, color: NAVY })],
    spacing: { after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Investment Thesis Memo — Rademikibart in Acute Exacerbations of Asthma & COPD", size: 26, color: GREY, italics: true })],
    spacing: { after: 60 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Working analyst file — prepared 20-Jul-2026. Companion valuation workbook: CNTB_Valuation_Model.xlsx.", size: 19, color: GREY })],
    spacing: { after: 60 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Not financial advice; not a recommendation to buy or sell any security. This document reflects one analyst's working assumptions, several of which are explicitly labeled ESTIMATES or UNKNOWNS. See the companion workbook for every underlying figure and source.", size: 18, color: GREY, italics: true })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF", space: 8 } },
    spacing: { after: 300 },
  }),
];

const section1to7 = [
  h1("1. Executive Summary"),
  p("Connect Biopharma is a single-asset, clinical-stage biotech built around rademikibart, an anti-IL-4Ra monoclonal antibody, now focused on a genuinely differentiated niche: treating acute exacerbations of asthma and COPD at the point of care (ED/urgent care), a setting where no biologic is approved today. Two Phase 2 studies (SEABREEZE STAT Asthma and COPD, ~160 patients each) are essentially fully enrolled; Asthma topline data are guided for early September 2026, with COPD data following shortly after."),
  p("The stock trades near, and at some points below, net cash - a reflection of a rocky company history (a shelved 2022 global Phase 3 program in atopic dermatitis, an ~80% stock decline in 2023, a narrowed pipeline, and new US-centric leadership in 2024) as much as of the two binary readouts still ahead. Our own bottom-up rNPV (Base case) is $5.07/share; blended with a thin, currently depressed peer-based relative valuation, we get to a Base blended target of ~$3.48/share - roughly 45% above the ~$2.40 current price, but well below the ~$7.00 Street consensus median, largely because (a) our relative-valuation leg is dragged down by a public peer (Upstream Bio) that itself just took a ~47% hit on a disappointing Phase 2 respiratory readout, and (b) our fully-diluted share count (~78.5mm, corrected against the actual Q1 2026 10-Q's option table) is meaningfully higher than the placeholder we used before we had that filing."),
  p("Management guides cash runway into 2H 2027; a simpler mechanical extrapolation of the current burn rate is tighter (roughly Q4 2026-Q1 2027). A financing event is likely regardless of what the data show, and highly likely if the data are positive (to fund Phase 3) - probably before investors get much time to enjoy a positive readout."),
  p("Bottom line: this is a binary, event-driven small-cap biotech position, not a steady compounder. Sizing and expectations should reflect that a negative readout on either program would remove the large majority of the modeled value, while a clean positive readout on both could be worth several multiples of the current price - see the Price Target Summary tab in the workbook for the explicit Low/Base/High spread ($0.74 / $5.07 / $39.34 per share on the rNPV alone)."),

  h1("2. The Long Thesis"),
  h2("Why this could work"),
  bullet("A validated mechanism in a new setting: IL-4Ra blockade is proven biology (Dupixent), and rademikibart's own prior data - Phase 1 IV (rapid FEV1 improvement within 24 hours, sustained for weeks) and an earlier global Phase 2 in chronic moderate-to-severe asthma (published in AJRCCM) - are directionally consistent with a fast-acting anti-inflammatory effect that could matter more in an acute-care setting than in chronic maintenance therapy."),
  bullet("A genuine white space: no biologic is approved for acute exacerbations of asthma or COPD today. If SEABREEZE STAT reads out clean, rademikibart could be first-to-market in a real unmet-need setting (the company cites over 1 million annual US ED visits for acute asthma exacerbation alone)."),
  bullet("A real, largely de-risked call option in China: Simcere's Phase 3 China atopic-dermatitis trial reportedly met all endpoints, and an NDA has been submitted - CNTB is entitled to up to ~$110mm in further milestones plus tiered low-double-digit royalties on Greater China sales, with none of the US clinical risk attached."),
  bullet("A sharper, more disciplined company than the one that IPO'd in 2021: the 2024 leadership reset narrowed the company to its single most differentiated shot rather than spreading capital across three unrelated candidates (AD/CRSwNP, ulcerative colitis, pruritus) as in 2021-2022."),
  h2("Why the market might be under- (or over-) pricing it right now"),
  p("The under-pricing case: the stock trades close to net cash ahead of two genuinely binary catalysts roughly six to ten weeks out. If the mechanism plays out as the supportive data suggest, the payoff structure is asymmetric - a modest downside (cash-ish floor) against a multi-billion-dollar peak-sales ceiling in a competition-free niche."),
  p("The more skeptical case: this is the same molecule's third attempted indication family (atopic dermatitis, then chronic asthma, now acute exacerbations specifically), following a 2022 decision to shelve the original global Phase 3 AD program and a subsequent ~80% stock decline. The specific endpoint being tested now - 28-day treatment-failure rate after a single dose in an acute exacerbation - has no biologic precedent for this mechanism; even Dupixent is not approved in this setting. Every pricing and adoption assumption in our model (and, we'd guess, in the sell-side models behind the current $4-10 price targets) is a genuine estimate, because this would be a new site-of-care commercial model with no real-world comparable."),

  h1("3. Origin Story & Corporate History"),
  p("Connect Biopharma was incorporated in November 2015 in the Cayman Islands, built around operations in China (Suzhou/Taicang) and a \u201cT-cell-driven\u201d immunology discovery platform, and IPO'd on Nasdaq in March 2021 at $18.49/share (total gross proceeds of $219.9mm before underwriting discounts) - confirmed directly from the Q1 2026 10-Q. At IPO, the pipeline had three candidates: CBP-201 (now rademikibart) for atopic dermatitis, asthma and chronic rhinosinusitis; CBP-307/icanbelimod, an S1P1 modulator for ulcerative colitis; and CBP-174, an H3-receptor antagonist for pruritus."),
  p("In November 2021, a global Phase 2 trial of CBP-201 in moderate-to-severe atopic dermatitis met its primary endpoint, and the company guided toward a global Phase 3 program starting mid-2022. In December 2022, however, the company announced it would not start that global Phase 3 program as planned, citing the \u201cchallenging funding environment,\u201d and would instead wait for a partner - a materially worse outcome than the original plan, delivered with little warning. Through 2023 the stock fell roughly 80% amid disappointing results tied to both lead candidates, and CBP-307 development was shelved entirely, narrowing the pipeline down to rademikibart alone."),
  note("Nuance worth holding onto: the stalled global strategy was distinct from the drug's China-specific data. Connect's China-only pivotal AD trial (run for the Simcere partnership) met all its primary and secondary endpoints around the same period. The 2022-2023 setback reads more like a funding/strategic retreat from a global, capital-intensive AD launch than a failure of the molecule's biology - but it was still a material, value-destroying reset that shareholders lived through, and it explains why the eventual Simcere deal looks like a monetization of a discontinued global strategy (all-China rights, milestone-plus-royalty) rather than a from-scratch partnership."),
  p("In 2024, a new, US-centric leadership team led by CEO Barry Quart, Pharm.D. took over and refocused the company on a single, narrower bet: rademikibart in acute exacerbations of asthma and COPD - a US-focused niche with no approved biologic competitor - while leaving the atopic-dermatitis opportunity to Simcere's China-licensed program. The SEABREEZE STAT Phase 2 studies were initiated in 2025, completed enrollment in mid-2026, and the company raised $20.2mm in a private placement in March 2026 to fund the run into data."),
  p("Investors should read this history as a genuine, twice-repeated pattern: rademikibart has now been re-pointed at three different indication families (AD, then chronic asthma, now acute exacerbations) as earlier plans didn't pan out as hoped. That could reflect smart optionality-seeking around a validated molecule with a new, disciplined management team - or a search for the indication where the data will finally read out cleanly. Both readings are consistent with the facts as we found them."),

  h1("4. Key Risks"),
  h2("4.1 Regulatory & Clinical"),
  bullet("Two genuinely binary Phase 2 readouts (Asthma, early Sept-2026; COPD, our estimate ~Q4-2026) on an endpoint - 28-day treatment-failure rate after a single dose in an acute exacerbation - with no biologic precedent for this mechanism. This is the single largest near-term risk to the entire investment case."),
  bullet("COPD carries extra regulatory risk on top of the shared clinical risk: the class-validating precedent (Dupixent) took years longer to reach COPD approval than asthma/AD, in a narrower population - suggesting FDA caution in COPD specifically, independent of rademikibart's own data."),
  bullet("Materialized risk, same franchise: this is not this molecule's first strategic reset - the December 2022 decision to shelve the global Phase 3 AD program is a direct, on-point precedent for execution/strategy risk in this exact program, not a generic industry risk."),
  bullet("No confirmed FDA alignment meeting on a Phase 3 design has occurred; management \u201cintends to move quickly\u201d post-data, but nothing is locked in, and timeline slippage is common industry-wide."),
  h2("4.2 Commercial & Market"),
  bullet("Zero pricing or reimbursement precedent for an ED/urgent-care-administered acute biologic - this would be a new site-of-care commercial model, not an extension of an existing one (unlike, e.g., a new chronic biologic entering an established specialty-pharmacy channel)."),
  bullet("The company's own cited peak-sales figures (>$3bn asthma, >$2bn COPD) come from a RedChip investor-relations research profile - a paid promotional/IR research service, not an independent sell-side or epidemiology source - and we could not corroborate them against independent third-party market sizing. Treat them as promotional, not as diligenced estimates."),
  bullet("Partner-execution risk is no longer hypothetical: per the actual Q1 2026 10-Q, of the original $123mm Simcere milestone pool (on top of a $21mm 2023 upfront), $5mm was received in 2024 but a further ~$8mm time-based milestone LAPSED in 2025 because Simcere missed its deadline. The $110mm remaining eligible balance is real, but its collection now has a documented, on-point precedent for slipping - not a generic third-party-dependency risk."),
  h2("4.3 Financial & Capital Structure"),
  bullet("$52.0mm total liquidity as of 31-Mar-2026 ($46.034mm cash and equivalents plus $5.997mm short-term investments - corrected from an earlier mis-split of a secondary source's headline figure) against roughly $16mm/quarter of operating burn. The 10-Q's own formal going-concern language commits only to funding being sufficient for at least one year from the 12-May-2026 filing date (i.e., through at least ~May-2027) - looser \u201cinto 2H 2027\u201d language appears only in press-release framing, not the filing itself. A financing event is likely around or shortly after the Phase 2 data regardless of outcome, and would be a distressed raise if the data disappoint."),
  bullet("Real, recent dilution, and a bigger overhang than headline share counts suggest: the March 2026 private placement alone added 6.13mm shares (~10% of the then-share count) at $3.25, and separately, the Q1 2026 10-Q shows 15.17mm stock options outstanding at a weighted-average exercise price of just $2.08 - meaningfully in-the-money at the current ~$2.40 share price, not a distant hypothetical. Fully-diluted share count (~78.5mm) runs well above the ~62.7mm basic count typically quoted."),
  bullet("Zero debt today is a genuine positive (no covenants, no maturity wall) - but it also reflects a company that has funded itself almost entirely through repeated equity dilution rather than any other structure."),
  h2("4.4 Ownership & Governance"),
  bullet("The 2024 CEO/leadership transition followed directly on the 2022-2023 strategic reset - worth tracking how much influence the founding team (Zheng Wei, Wu Bin Pan) retains at the board level going forward; we did not find a detailed current board composition in this search."),
  bullet("A related-party financing, not generic insider buying: per the Q1 2026 10-Q (Note 9), the March 2026 private placement was led by Panacea Venture, described in the filing as the Company's largest current investor, which purchased $4.0mm of the $20.2mm round. Board member James Huang is the sole owner of Panacea Innovation Limited, which is the sole owner of Panacea Venture - meaning the company's own largest shareholder, controlled by a sitting director, anchored its own financing round. Read this as a genuinely stronger capital-commitment signal than ordinary PIPE participation (the largest holder underwriting the raise) but also as a governance question worth asking directly: were the $3.25/share terms arm's-length, and how was the related-party transaction approved? Separately, James Huang's ~$372,000 open-market purchase in May 2026 (150,000 shares at $2.48) is a distinct, smaller, genuinely open-market signal and should not be conflated with the PIPE-anchoring role above."),

  h1("5. Balance Sheet & Cash Runway Summary"),
  p("As of the actual Q1 2026 10-Q balance sheet (quarter ended 31-Mar-2026): cash and cash equivalents were $46.034mm, plus a separate $5.997mm of short-term investments, for total liquidity of $52.03mm - about $6mm more than the headline \u201c$46.0mm\u201d figure widely quoted in press summaries, which turns out to describe only the cash-and-equivalents line, not the total. Debt is zero. The March 2026 private placement (6.13mm shares at $3.25, ~$20.2mm gross / ~$18.6mm net) closed on the last day of the quarter and is included in the balance; it was led by Panacea Venture, the company's largest shareholder (see Section 6). Q1 2026 R&D expense was $15.03mm (more than double the prior-year quarter, reflecting SEABREEZE STAT enrollment), G&A was $4.75mm, and license/collaboration revenue from Simcere was a token $0.17mm (amortization of the 2023 upfront fee, not new cash). Net loss widened to $19.40mm ($0.34/share) from $10.27mm ($0.19/share) a year earlier; operating cash outflow was $16.0mm for the quarter."),
  p("The 10-Q's own formal Liquidity & Going Concern note (Note 2) commits only to cash being sufficient for \u201cat least one year from the date this Quarterly Report is filed\u201d - filed 12-May-2026, i.e., a floor of roughly May-2027. The looser \u201cinto 2H 2027\u201d framing appears in press-release/8-K language, not in this more conservative formal disclosure, and a simple mechanical extrapolation of the current burn rate against current cash points to a similar, maybe slightly tighter, window. We would treat ~May-2027 as the company's actual commitment and \u201c2H 2027\u201d as aspirational guidance layered on top of it, and would not assume the longer figure without an updated Q2 2026 10-Q (due ~12-Aug-2026) confirming the trend. See the Balance Sheet & Cash Runway tab in the workbook for the full build and the explicit source citations."),

  h1("6. Institutional Ownership & Insider Activity"),
  p("Per the most recent aggregation of 13D/G and 13F filings we located, institutions hold roughly 59% of CNTB's shares (28 institutional filers, ~8.3mm shares in that count). Named holders include Perceptive Advisors (a dedicated healthcare/biotech long-only specialist - arguably the most thesis-relevant holder on the list), Ikarian Capital, Boothbay Fund Management, BML Capital Management, Renaissance Technologies, Citadel Advisors, and Jane Street Group. The presence of Renaissance, Citadel and Jane Street should be read as largely quantitative/market-making flow rather than a directional fundamental view on the Phase 2 data - don't over-read their presence as smart-money conviction."),
  p("On insiders: as detailed in the risk section above, distinguish James Huang's personal open-market purchase (real conviction) from Panacea Venture's role anchoring the March 2026 PIPE as the company's largest shareholder (a related-party financing, not a market-signal purchase in the same sense). We do not have Panacea Venture's total stake size from the 10-Q alone - a Schedule 13D/G pull is the natural next step, since it is arguably the single most important ownership fact in this file and was not fully captured in our earlier institutional-holder research. We did not find other Section 16 officer/director transactions of note in this search - worth pulling a fresh Form 4 screen closer to the September data."),

  h1("7. Where Outside Expertise Would Help Most"),
  p("Ranked by how much a real conversation would likely beat what's available in public filings:"),
  bullet("1. ED / pulmonary / critical-care physicians: is there genuine clinical and workflow appetite to administer a new biologic in the acute ED/urgent-care setting? This is the single biggest unknown driving our peak-penetration assumption, and nothing in the public record addresses it directly."),
  bullet("2. Payer / managed-care contacts: how would a single-dose, acute-episode biologic actually get coded and reimbursed - buy-and-bill, specialty pharmacy, or bundled into a hospital DRG? This affects realized net price more than almost any other input in our model."),
  bullet("3. Former Connect Biopharma or Simcere employees, or SEABREEZE trial site staff: color on why the original global AD Phase 3 program was really shelved in December 2022 (funding vs. strategic vs. data-driven) would sharpen our read of that materialized risk."),
  bullet("4. Pulmonology KOLs close to the SEABREEZE STAT trial sites: any directional read on enrollment characteristics, control-arm treatment-failure rates, or biomarker cut-points ahead of the formal topline (with appropriate care around trial-integrity and MNPI concerns)."),
  bullet("5. Biotech-focused bankers: whether Simcere or any other party has expressed interest in ex-China rights, especially in light of the recent Apogee/AbbVie and Aiolos/GSK transactions in the same IL-4/IL-13/TSLP respiratory space."),
];

module.exports = { h1, h2, p, bullet, note, cell, numbering, titleBlock, section1to7, NAVY, GREY, LIGHTBLUE };

const { Document, Packer, Paragraph, TextRun, Table, TableRow, WidthType, AlignmentType, BorderStyle } = require("docx");
const fs = require("fs");
const { h1, h2, p, bullet, note, cell, numbering, titleBlock, section1to7, NAVY, GREY, LIGHTBLUE } = require("./memo_partA.js");

function valTable() {
  const headerRow = new TableRow({
    children: [
      cell("Method", { bold: true, shade: NAVY, color: "FFFFFF", width: 34 }),
      cell("Value / share", { bold: true, shade: NAVY, color: "FFFFFF", width: 22, align: AlignmentType.CENTER }),
      cell("Role in blend", { bold: true, shade: NAVY, color: "FFFFFF", width: 44 }),
    ],
  });
  const rows = [
    ["rNPV - bottom-up, risk-adjusted (Base)", "$5.07", "60% weight - primary method; most conservative/defensible number in the file"],
    ["Relative valuation - EV / risk-adj. peak sales (vs. Upstream Bio)", "$1.08", "40% weight - thin (one true peer), currently depressed by UPB's own post-miss selloff"],
    ["Blended price target (Base case, ~12mo)", "$3.48", "Primary output - see Price Target Summary tab"],
    ["rNPV - Low scenario", "$0.74", "Reference only - both readouts disappoint or execution stalls"],
    ["rNPV - High scenario", "$39.34", "Reference only - compounded best-case across every input; not a realistic point forecast"],
    ["DCF - unrisked \u201cif it succeeds\u201d ceiling case", "$16.73", "Reference only - explicitly NOT risk-adjusted, not blended into the target"],
    ["Current share price (~1-Jul-2026)", "$2.40", "Highly volatile - 52-week range $1.23-$3.82"],
    ["Street consensus target (7 analysts, Jul-2026)", "$7.00", "Median of Buy/Overweight/Outperform ratings - see Section 9"],
  ];
  const dataRows = rows.map((r, i) => new TableRow({
    children: [
      cell(r[0], { shade: i % 2 === 0 ? "F2F2F2" : "FFFFFF", width: 34 }),
      cell(r[1], { shade: i % 2 === 0 ? "F2F2F2" : "FFFFFF", width: 22, align: AlignmentType.CENTER, bold: i === 2 }),
      cell(r[2], { shade: i % 2 === 0 ? "F2F2F2" : "FFFFFF", width: 44 }),
    ],
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}

const section8to10 = [
  ...[],
  valTable(),
  new Paragraph({ text: "", spacing: { after: 200 } }),
  p("Every dollar in the table above is fully sourced and formula-driven in the companion workbook, with Low/Base/High assumptions and a single scenario selector that flows through the rNPV, DCF and Relative Valuation tabs. The Base-case blend deliberately underweights the DCF ceiling case (excluded entirely) because including an unrisked \u201cif it succeeds\u201d number in a probability-weighted target would double-count the upside that the rNPV tab already prices in properly."),
  p("Note on a pending update: the Loss-of-exclusivity assumption (Cover & Assumptions, row 52) was revised in July 2026 from an unverified placeholder (2038/2040/2042) to figures sourced from the FY2024 and FY2025 10-Ks' actual patent-expiry disclosures (2037/2040/2045 - see Section 4.1 and the workbook's source note on that row). The workbook has not yet been reopened in Excel to recalculate the downstream rNPV/DCF outputs against this change, so the Low/Base/High figures in the table above still reflect the prior assumption until that recalc happens - treat them as pending, not final, on this one input."),

  h1("9. Sell-Side & Third-Party Research Reconciliation"),
  p("As of mid-July 2026, seven sell-side firms cover CNTB, all bullish: H.C. Wainwright ($7, Buy), BTIG ($10, Buy), Cantor Fitzgerald ($4, Overweight), Canaccord Genuity ($6, Buy), Piper Sandler ($7, Overweight/Strong-Buy), and Oppenheimer ($8, Outperform, initiated 9-Jul-2026), for a consensus median around $7.00 and a \u201cStrong Buy\u201d aggregate rating."),
  p("A pattern worth naming: essentially all of this coverage was initiated in the roughly ten weeks before the September 2026 data (May-July 2026) - classic sell-side positioning ahead of a binary catalyst. None of these firms has a track record of being right or wrong on an actual CNTB data readout yet; their price targets are, like ours, pre-data estimates built on similar uncertain pricing and adoption assumptions, not a verified consensus."),
  p("Conflicts of interest: now confirmed, not merely suspected. The 8-K announcing the March 2026 private placement names Leerink Partners LLC and Cantor Fitzgerald & Co. as joint placement agents. Cantor Fitzgerald carries a live $4 Overweight rating on CNTB in the table above - a direct conflict between its role arranging a related-party-anchored financing and its ongoing sell-side coverage. This doesn't necessarily mean the rating is wrong, but it should be weighted accordingly, and it's a genuine data point this file lacked until now."),
  p("Where our numbers sit relative to the Street: our rNPV Base case ($5.07, using the fully-diluted share count confirmed in the actual Q1 2026 10-Q) sits inside the low-to-high range of sell-side targets ($4-$10), but near the low end. Our blended target ($3.48) sits below the entire Street range, because our Relative Valuation leg is dragged down by Upstream Bio's own currently-depressed multiple following its Phase 2 miss - strip that leg out and our own bottom-up view is much closer to consensus. We read this as the Street generally not yet reflecting a peer's bad-news read-through (and possibly not yet reflecting the full ~78.5mm fully-diluted share count either), rather than as evidence our rNPV is too conservative."),
  p("On named public calls: beyond the sell-side notes above, the only independent (non-institutional) public commentary we located was a small paid/independent newsletter (a Substack post from a micro-cap-focused writer) framing the 2024 strategic reset and cash runway positively. We found no specific, sizeable, named-investor or well-known-YouTuber call on CNTB in this search - if one exists, it wasn't prominent in what we searched."),

  h1("10. Independent Scrutiny of Key Data"),
  p("We do not have access to patient-level data or full summary statistics for the SEABREEZE STAT Asthma or COPD studies - they have not yet reported, so there is nothing to independently recompute on the pivotal readouts themselves. That is the whole point of the September/Q4 2026 catalysts."),
  p("On the supportive prior data, we now have real numbers rather than press-release characterizations for the Phase 1 IV study: a March 2026 topline release (study CBP-201-105) reported mean FEV1 improvements of ~200-400 mL maintained through Day 29 in both asthma and COPD patients (vs. placebo trending down over the same period), with clinically meaningful increases of 100-400 mL observed in many patients as early as 15 minutes post-dosing. This was a small, single-dose, placebo-controlled pharmacology study (12 asthma and 10 COPD patients total, randomized 4:1 drug:placebo) - genuinely supportive mechanistic signal, but not remotely powered for efficacy, and shouldn't be over-weighted relative to the much larger SEABREEZE STAT readouts. The earlier global Phase 2 published in AJRCCM remains unpulled in full - we still have only the topline characterization for that one, and would recommend reading the actual publication before giving it more than directional weight."),
  p("The company's cited peak-sales claims (>$3bn asthma, >$2bn COPD) trace to a RedChip investor-relations research profile. RedChip is a paid promotional/IR research provider, not an independent equity-research house or epidemiology consultancy, and we could not find the underlying independent market-sizing study it claims to summarize. Our own bottom-up patient-population and pricing build in the companion workbook is deliberately independent of this figure, and comes out lower - we'd treat the RedChip numbers as marketing collateral, not diligence."),
  p("Simcere's China Phase 3 atopic-dermatitis trial now has a name and real disclosed numbers: RADIANT-AD, a 259-patient, double-blind, placebo-controlled study, with 52-week topline data presented at the March 2026 American Academy of Dermatology (AAD) annual meeting (a company 8-K exhibit, not yet a peer-reviewed publication or NMPA-level data package). Actual results: 96.6% of patients achieved EASI-75, 87.1% achieved IGA 0/1 (clear or almost clear), and 85.3% achieved EASI-90 at Week 52; safety was reported as comparable to placebo with low conjunctivitis rates. The previously-flagged \u201cnear-maximal responses in ~90% of patients\u201d characterization holds up reasonably well as a rough average of these three real endpoints - a fair paraphrase, not marketing inflation, though it's still a company-sourced topline rather than an independently adjudicated dataset. One new, currently unmodeled data point: CEO Quart's remarks alongside this release explicitly floated ex-China AD as a future indication \u201cfor us or a future partner\u201d - not in the companion workbook or blended price target anywhere today, and worth flagging as optionality rather than base-case value."),

  new Paragraph({ text: "", spacing: { after: 100 } }),
  note("This memo and its companion workbook are a working file. Please send follow-up documents, screenshots, or questions as they come in — including the Q2 2026 10-Q (due ~12-Aug-2026) and the SEABREEZE STAT Asthma topline (expected early September 2026) — and we will reconcile them against everything above, updating and flagging any changes or corrections explicitly."),
];

const doc = new Document({
  numbering,
  styles: {
    default: {
      document: { run: { font: "Arial", size: 21 } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { bold: true, size: 28, color: NAVY, font: "Arial" }, paragraph: { spacing: { before: 320, after: 160 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { bold: true, size: 24, color: "2E5395", font: "Arial" }, paragraph: { spacing: { before: 240, after: 120 } } },
    ],
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1260, right: 1260 } },
    },
    children: [
      ...titleBlock,
      ...section1to7,
      ...section8to10,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + "/../CNTB_Investment_Thesis_Memo.docx", buf);
  console.log("written");
});

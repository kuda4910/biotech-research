# Thesis Update / News Reconciliation Template

Use this whenever you have a new piece of information about a company already tracked in this
repo — a filing, a press release, a news article, an earnings call, a screenshot, a forum post, an
analyst note — and want it reconciled against the existing thesis rather than just discussed in
the abstract. Paste this as your message (in a fresh chat, or right here in Claude Code) along with
the new material, filling in the ticker.

---

I have new information about **[TICKER]**, already tracked in this repo at `companies/[TICKER]/`.
Reconcile it against the existing thesis rather than treating it as a standalone question. Here's
the material: **[paste the article/filing text, attach the file, or describe what you saw/heard]**

## Process to follow

1. **Load current state first.** Read, in this order: the `[TICKER]` section of `README.md`, the
   `[TICKER]` entry in `portfolio.json`, and `companies/[TICKER]/[TICKER]_Project_Context.md` if it
   exists (CNTB has one; treat it as the fuller backing detail behind the README summary). If you
   need the thesis memo's actual text (not just the README summary), extract it — `pandoc -t plain
   file.docx` or `python-docx` if pandoc isn't available — rather than guessing at what it says.

2. **Compare, don't just summarize.** For each fact in the new material, check it against what's
   already on file and sort it into one of:
   - **Confirms** something already assumed — note it, no edit needed.
   - **Contradicts or corrects** something already stated as fact — this is the important case.
     Say explicitly what was wrong, what's right now, and where the wrong version came from (was it
     a placeholder guess, a misread secondary source, stale data?). Match the tone of the existing
     corrections logs (see README's "Known corrections already made" sections) — own it plainly,
     don't bury it in a vague rewrite.
   - **Resolves an open question** — check it off, cite the source.
   - **Raises a genuinely new risk or question** — add it, ranked by how much it could move the
     valuation, not just added to the bottom of a list.
   - **Moves a model assumption** — a probability-of-success input, a patient-population estimate,
     a pricing assumption, a milestone/royalty term, a share count, a cash balance. This is the
     case that actually changes numbers, not just narrative.

3. **Surface it before touching anything.** Present what you found — confirms/corrections/new
   risks/assumption changes — and your read on materiality, and ask before editing any file. Don't
   silently move a bear/mid/bull number or rewrite the thesis on your own judgment call; that's the
   user's decision to make, informed by your analysis. Never inflate the bull case or manufacture
   bearishness to make the update feel more dramatic than it is — if the news is genuinely minor,
   say so.

4. **On confirmation, make the edits — and keep every file in sync:**
   - `portfolio.json` — update `thesis`, `price` (only if you have a real current quote, not a
     guess), `bear`/`mid`/`bull` (only if a model assumption actually changed and you've recomputed
     the model's own Low/Base/High or blended output — don't hand-edit these independent of the
     model), `catalyst`/`catdate`, `questions`, and this position's `updated` field, plus the
     top-level `updated` field.
   - `README.md`'s section for this company — update the relevant facts, add to "known corrections
     already made" or the open-questions list as appropriate. Keep the same voice/density as the
     existing SPRB/CNTB sections.
   - The valuation model (`.xlsx`) — if an assumption changed, edit the actual Low/Base/High input
     cell in Cover & Assumptions (via `openpyxl` — don't hand-calculate downstream tabs, they're
     formula-driven and will update themselves). If a build-script toolchain exists for this
     company (check for a `build/` folder, as CNTB has), prefer adding a new `patchN.py` over
     editing the saved workbook directly, so the change is reproducible and logged. Recalculate
     and confirm zero formula errors before considering this done.
   - The thesis memo (`.docx`) — update the affected section(s) so the memo's numbers and narrative
     match the model. If a `memo_part*.js` build toolchain exists, prefer editing that and
     rebuilding over hand-patching the binary.
   - Log the correction plainly — a one-line addition to whichever "corrections log" already
     exists for this company (README section and/or `*_Project_Context.md`), not a silent rewrite.

5. **Verify consistency before committing.** Grep the memo and README for any old figure you just
   changed in the model — this exact failure mode (a number fixed in one place but stale in
   another) has happened before in this repo; don't repeat it.

6. **Commit and push per the repo's established workflow** — direct commits to `main`, no PR step
   (see `README.md`). Write a commit message that names the source of the update and what changed,
   the same way the CNTB 10-Q correction commits would read.

If the material actually describes a **new company** not yet tracked here, stop and use
`Biotech_DD_Prompt_Template.md` instead — this template is for updating an existing thesis, not
starting one.

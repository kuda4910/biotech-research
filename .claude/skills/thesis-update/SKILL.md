---
name: thesis-update
description: Reconcile new news, filings, or research about an already-tracked biotech position (SPRB, CNTB, or any future company in companies/) against its existing thesis, model, and README section, then update all affected files in sync. Use when the user pastes/attaches news, a filing, an article, or a claim about a company that already has a folder under companies/ — not for starting a brand-new company (that's the DD template).
---

Read `templates/News_Update_Prompt_Template.md` in this repo root and follow its process exactly,
using whatever news/document/claim the user just supplied in this conversation as the input
material. That file is the single source of truth for this workflow (file-sync rules, tone,
commit conventions) — don't duplicate or improvise a different process here.

Quick reminders that matter most in practice:
- Load the company's current state (README section, `portfolio.json` entry, and
  `companies/<TICKER>/<TICKER>_Project_Context.md` if present) before comparing anything.
- Surface confirms/corrections/new-risks/assumption-changes and get confirmation before editing —
  never silently move a bear/mid/bull number or rewrite thesis text on your own judgment.
- If a model assumption actually changes, edit the Low/Base/High input cell in the workbook (via
  `openpyxl`, preferring a new `patchN.py` if a `build/` toolchain exists) and let the formulas
  flow through — don't hand-edit downstream numbers.
- Keep `portfolio.json`, `README.md`, the thesis memo, and the model in sync — this repo has
  already had the "fixed in one place, stale in another" bug happen once; don't repeat it.
- Commit and push straight to `main` per the repo's established direct-commit workflow, with a
  message naming the source of the update.
- If the material describes a company not yet tracked here, redirect to
  `templates/Biotech_DD_Prompt_Template.md` instead.

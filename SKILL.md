---
name: deliverable-consistency-gate
description: Runs a consistency gate over a client deliverable set — deck, memo, workbook — and reports where the same claim carries two different numbers, where a total row does not equal its parts, where shares miss 100%, where a cited exhibit number does not exist, and where a TODO survives into the file. Reads .pptx, .docx and .xlsx without third-party libraries. Use when a deck and a workbook go out together, when a board pack or an investor memo is about to be sent, when someone asks whether the numbers agree across the files, or when a deliverable must pass a partner review. Not for reformatting slides, rewriting the narrative, or auditing the underlying data model.
---

# Deliverable Consistency Gate

A deliverable is not one file. A deck says $12.4M, the memo says $13.1M, the workbook totals a column that does not add up, and the deck cites Figure 5 of a set that contains three figures. Each file looks fine alone; the set contradicts itself. This skill reads the set, extracts every number with its label and unit, and reports the contradictions with a locator for each side.

## Contract

1. **Cross-file, not single-file.** A number is only checked against the same claim elsewhere: deck, memo, workbook, appendix. A single file cannot contradict itself invisibly.
2. **Normalise before comparing.** `$12.4M`, `$12,400K` and `12.4 million` are the same value. Compare in base units, then report the original strings.
3. **Rounding is not a conflict.** A 0.4% difference between a rounded slide figure and a workbook cell is expected. The tolerance comes from the printed precision.
4. **A locator on both sides.** `slide 3 · para 2` and `Revenue!C5`. A finding without two locators is not a finding.
5. **Blockers are arithmetic and leftovers.** A total row that does not add up, a surviving TODO: these cannot be argued with. A label conflict may be legitimate (constant currency, a restated segment) and is a warning that needs a human sentence.
6. **Never silently "fix" the files.** The gate reports; the owner decides which side is right.

## Workflow

### Step 0 — Assemble the set

Collect the files that go out together: `.pptx`, `.docx`, `.xlsx`, plus text appendices. Pass the folder or list the files:

```bash
python scripts/consistency_gate.py --input board-pack/ --out out/
python scripts/consistency_gate.py --input deck.pptx --input memo.docx --input model.xlsx --out out/
```

Temporary files (`~$...`) are skipped. No libraries required: the script unzips and parses the OOXML directly, so it runs on a bare Python install.

### Step 1 — Read the verdict

- `BLOCK` — a total does not add up, or a TODO/lorem survives. Do not send.
- `WARN` — the same label carries two values, shares miss 100%, a cited exhibit does not exist, or a heading repeats.
- `PASS` — nothing above tolerance. Say so and say how many numbers were compared; a pass with 12 numbers is weaker evidence than a pass with 400.

### Step 2 — Triage the blockers

1. `TOTAL_MISMATCH`: get the row range from the locator, then decide whether the total or a line item is wrong. This is the one finding a reader will check first, usually on screen, in front of the client.
2. `PLACEHOLDER`: find and finish, or delete the sentence.
3. `PCT_TOTAL`: a share column at 98.5% usually means an "other" bucket was dropped, not a rounding issue.

### Step 3 — Triage the warnings

- `NUMBER_CONFLICT`: read both sentences. Constant-currency, restated segments and "plan vs actual" are legitimate; an unexplained 5% gap is not.
- `FIGURE_REF` / `TABLE_REF`: count the actual exhibits, then fix the numbers in the text. This is the cheapest credibility loss in a deck.
- `DUPLICATE_HEADING`: usually a merged section that was not merged.
- `DF`-style tolerance questions: state the tolerance used and why.

### Step 4 — Write the fix list

One line per finding, in the owner's language: "Revenue!C5 total is 2540; the three rows above sum to 3090 — either Partner (700) or the total is wrong." Send the JSON when the owner wants to work through it in a spreadsheet.

### Step 5 — Re-run after the edits

The gate is cheap. Re-run after the fixes and keep both reports; the second one is the artifact that says the set was checked before it went out.

## Output format

- Verdict line with file count, numbers parsed, exhibit counts.
- Findings table: severity, code, where, one sentence of fact with both values.
- Evidence section: each finding with its two locators and the surrounding text.
- Rule set: the six checks, so a reader knows what was not checked
- JSON: `consistency_gate.json` with every finding and its locators.

## Boundaries

- Does not open the files in Office, render them, or judge layout, typos, or visual hierarchy.
- Does not verify that the numbers are true — only that the set agrees with itself.
- Does not parse legacy `.doc`, `.ppt`, `.xls`, or PDF; convert first and say so.
- Does not reconcile definitions (that is metric reconciliation); this gate compares printed claims, not models.
- Does not edit files. It produces a fix list; the owner edits.

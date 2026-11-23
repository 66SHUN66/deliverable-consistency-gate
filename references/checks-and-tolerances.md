# Checks, tolerances, and false positives

## The six checks

| Code | Severity | What it proves | What it cannot see |
|---|---|---|---|
| `NUMBER_CONFLICT` | warning | the same label carries two values beyond rounding tolerance | whether one is constant-currency, plan, or restated |
| `TOTAL_MISMATCH` | blocker | a row labelled total differs from the contiguous rows above it | whether a row is intentionally excluded, or a sign convention |
| `PCT_TOTAL` | warning | a share column does not sum to 100% | whether an "other" bucket is implied |
| `FIGURE_REF` | warning | text cites a figure number the set does not contain | whether the figure lives in an appendix that was not passed in |
| `TABLE_REF` | warning | text cites a table number the set does not contain | same |
| `PLACEHOLDER` | blocker | TODO / TBD / lorem / `[insert]` survived | nothing — this one is always real |
| `DUPLICATE_HEADING` | note | a heading repeats inside one document | whether the repeat is intentional (running header) |

## Normalisation

Values are compared in base units:

| Written | Base |
|---|---|
| `$12.4M`, `12.4 million`, `$12,400K` | 12,400,000 money |
| `62%`, `6200 bps` | percent (bps divided by 100) |
| `1,450` with no unit | count |

The label key is the last three non-stopword tokens before the number: `Q3 revenue reached $12.4M` → `q3 revenue`. Verbs (`reached`, `was`, `grew`) and articles are dropped so that two differently phrased sentences about the same metric land in the same bucket.

## Tolerance

`tol = max(0.005, 0.5 × 10^-min_decimals)`, applied to the relative difference.

- `$12.4M` vs `$12,400K`: identical after normalisation, no finding.
- `62%` vs `61.8%`: 0.3% relative, inside tolerance for two-decimal reporting, no finding.
- `$12.4M` vs `$13.1M`: 5.3%, finding.

State the tolerance in the report whenever a finding is a warning. A reviewer who can see the tolerance can argue with it, which is the point.

## Known false positives

1. **Plan vs actual.** "Revenue was $13.1M, ahead of the plan of $11.8M" produces a conflict if both are labelled `revenue`. The label key does not know `plan`. Expected; the report is a list of claims for a human, not a verdict on arithmetic.
2. **Constant currency.** Same metric, same period, two bases: legitimate, and the sentence should say so.
3. **Segment vs total.** `Enterprise revenue 6.2` and `Revenue 12.4` differ by design. Only identical label keys are compared; a segment row usually carries `enterprise` in the key.
4. **Year and quarter numbers.** Four-digit years and `2026`-style ordinals are excluded; a bare `Q3` is not, but it has no unit and is dropped unless a currency or percent sign is attached.
5. **Template text.** A deadline table copied from last quarter produces conflicts the moment the period changes. That is a real finding, not a false positive.

## Total rows

The check reads a row whose first populated cell contains `total`, `totals`, `sum`, `subtotal` or `grand total` (case-insensitive), then sums the contiguous numeric rows directly above it in the same column, up to five rows back. It stops at the first gap.

- Multi-level totals (subtotal per segment, then a grand total) can trip this: the grand total sums segments, not the rows immediately above.
- A total that includes a row with a blank value will disagree; blanks are not zero in this check.
- Sign conventions (a negative refund line) are summed as written, which is correct for a P&L and wrong for a mix table.

## Reporting a conflict well

Give both locators, both strings, and the relative difference:

> `Q3 revenue` — deck slide 2 says `$12.4M`; memo para 2 says `$13.1M` (5.3%). Constant-currency or a restatement? One sentence in the memo settles it.

Do not write "the deck is wrong". The gate cannot know which side is right, and the person who does know is the one reading the report.

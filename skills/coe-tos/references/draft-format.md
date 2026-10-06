# Draft JSON format (version 1)

The host LLM analyzes the sources and writes this JSON. The CLI checks arithmetic
and preservation deterministically; it does not call an LLM or infer source meaning.

## Required top-level fields

| Field | Meaning |
|:--|:--|
| `version` | `1` |
| `mode` | `existing-exam` or `proposed-blueprint` |
| `total_points` | Positive numeric total; decimal scores are supported |
| `metadata` | Course/program/period/preparer values; see template field map |
| `topics` | Ordered `{id, title}` objects; order determines template rows |
| `questions` | `{id, points, criteria}` objects |
| `allocations` | Atomic scoring ledger entries |
| `assumptions` | Optional array of disclosed assumptions |
| `proposed_changes` | Optional array of explicit assessment changes |

Each criterion is `{id, points}` with optional `title` and `subcriteria`, a map of
published subcriterion ID to points. Questions may also carry a `title`. Include the
`subcriteria` map whenever the source publishes a scoring breakdown; excluding it
weakens validation. If the source says 5 verification points and 3 convergence
points, a 4/4 mapping must fail even though the combined score is still 8. Titles are
cosmetic: the generated review prints the ID when a title is absent, so an existing
draft without titles stays valid.

Each allocation has:

```json
{
  "id": "q1-algorithm-recall",
  "question": "Q1",
  "criterion": "C",
  "subcriterion": "recall",
  "topic": "bisection",
  "points": 2,
  "scores": {"remembering": 2, "understanding": 0, "thinking": 0},
  "evidence": [{"source": "exam.pdf", "locator": "page 6, rubric C recall"}],
  "rationale": "The source explicitly awards two points for recalling the update rule."
}
```

- IDs are unique strings within their collection; criterion IDs are unique per question.
- `question`, `criterion`, `subcriterion` and `topic` must reference declared entries.
- Scores are finite, nonnegative numbers; booleans are not scores.
- Each allocation's R+U+T equals its points.
- Entries per declared subcriterion/criterion/question sum to that published total.
- The total allocated points equal the declared exam total.
- Every allocation needs evidence and a rationale. Tool validation confirms their
  presence, not that a citation is semantically correct. The agent must inspect it.

## Metadata

The default profile uses `program`, `course_code`, `course_title`,
`academic_year`, `semester`, `assessment`, `preparer`. Additional metadata
(e.g. `program_full`, reviewer, learning-outcome definitions, instruction hours)
is retained in draft/review files. Add a field address to the profile to print
it in a template that supports that field. Do not invent teaching hours or
approval signatures. Missing mapped values become `[To be supplied]`; a field
with a restrictive list may require a valid selection before building.

### Input provenance

Record how the sources arrived so the review can show it:

| Field | Meaning |
|:--|:--|
| `input_mode` | `file`, `chat-pasted` or `chat-attachment`. `chat-pasted` adds a validation warning that page-image verification was impossible. |
| `sources` | Optional list of supplied file names or chat input labels. |
| `totals_confirmed_by` / `totals_confirmed_on` | Who confirmed the published rubric totals, and when. Expected for chat-supplied input. |

These fields are review metadata only. They never change printed template values
unless the profile maps them, and they do not affect scoring validation.

Default delivery retains metadata on Notes and every allocation's source names,
locators and rationale on Allocation ledger. draft.json is an internal build input,
not a required separate deliverable; external copies require --diagnostics.

## Teaching-hours allocation for proposed blueprints

When requested, derive proposed topic weights from actual hours:
`weight = topic hours / total hours`. Allocate integer question counts using
largest remainders so the specified total is retained. Store the actual
hours, calculated weights and rounding in metadata/assumptions and cite their
source. Convert that proposal into explicit scored tasks and ledger entries.
Do not use teaching hours to overwrite an existing exam's rubric scores.

## Template profile

Copy `assets/cjc-profile.json` and change the inspected field/column addresses.
All column keys from that example are required by the current renderer.
Templates without formulas in summary cells receive calculated values there;
original formulas are retained and their supported caches recomputed instead.
Percentages are stored as fractions (0.16 for 16%), retaining the template's
number format. Supply a template already formatted for percentages.

`bands` may be omitted for institutions without specified bands. `print_area`
and `fit_to_page` are explicit export settings; omit/disable them to retain
another template's print setup. The renderer does not insert rows, rebuild
the form, or automatically convert six-level Bloom layouts to three groups.

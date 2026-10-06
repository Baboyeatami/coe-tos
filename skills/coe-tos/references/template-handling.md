# Template preservation and native PDF export

## Strategy

An XLSX is an OOXML ZIP package. `workbooks.py` reads the supplied workbook,
changes cell values and supported formula caches in worksheet XML, and updates
only the profile's print area, fit-to-page properties and calculation flags.
It writes a staged package and checks fidelity before replacing the output.
All other parts retain their original bytes, including styles, grouped drawings,
images, VML comments, themes, relationships, content types and document metadata.

No load-and-save round trip through openpyxl is used to generate the artifact.
That avoids the grouped-shape/header-loss issue found in the original workflow.
openpyxl is used for reading and in-memory calculation only.

## Fidelity boundaries

The gate checks all package names, byte-identical non-edited parts, unchanged
worksheet structure outside declared value/cache/print edits, cell style
attributes and original formulas. It does not certify pixel-identical rendering
across Office applications. Fonts, printer metrics and export engines differ.
Changing a supplied template's design requires explicit template editing outside
the preservation workflow, followed by a new inspected profile.

The renderer accepts XLSX templates with a contiguous three-group topic grid.
Macro-enabled input may be extracted, but generation requires an XLSX template.
Digital signatures cannot remain valid after edits; signed/encrypted forms
need a suitable editable template. No workbook macros are executed by the scripts.

## Supplementary worksheets

`--with-mapping` appends generated worksheets (`Assessment mapping`, `Topic
allocation`, `Allocation ledger`, `Notes`) to the delivered workbook. They are
written straight into the OOXML package: a new worksheet part plus the matching
entries in `xl/workbook.xml`, `xl/_rels/workbook.xml.rels` and
`[Content_Types].xml`. No library save round trip occurs, so the template's own
worksheet, styles, drawings and media stay byte-identical.

The new sheets use inline strings and no formulas, so nothing needs
recalculating. Percentages are written as preformatted text because the
template's style table is deliberately left untouched.

The fidelity gate treats these additions as declared: `verify()` requires the
set of added parts to match `fidelity.json` exactly, removes only the declared
sheet entries and relationships before comparing, and still fails on any other
package, worksheet or workbook change. An undeclared sheet is refused.

They are hidden by default because Excel does not print hidden sheets, which
keeps the exported form at one page. `--mapping-visible` prints them, and the
tables can split across pages; review every page if you use it.

Once a sheet is appended, the workbook is no longer byte-identical to the
institutional original. Use this only when the reader asks for the mapping inside
Excel.

## Test Item No. cells

The institutional row height is fixed, so a topic that maps to several criteria
can overflow its `Test Item No.` cell. Set `refs_max_chars` in the profile to the
usable width (the bundled CJC profile uses 16). References are then abbreviated as
`Q1-c1b +2`: as many as fit, plus the number omitted. No information is lost,
because the full list appears in the allocation ledger, in the generated
`Topic allocation` sheet's Refs column, and in `review.md`. Omit `refs_max_chars`
to print every reference in full.

## Formulas

The deterministic evaluator supports cell/range references, cross-sheet references,
arithmetic, comparisons, IF, SUM, AND and OR, covering the bundled template.
Other formula caches are cleared and listed in `fidelity.json`, with native
recalculation enabled. The formula XML remains intact. Formula evaluation has
no Python `eval` or execution of worksheet content.

## Print and visual checks

`excel-mac` drives an installed Microsoft Excel through AppleScript on a temporary
copy and closes only the workbook it opened. macOS may request automation
permission. `libreoffice` uses an isolated temporary user profile and native Calc
PDF export. Neither renderer saves back over the verified workbook.

The supplied CJC template keeps its original page size/orientation. The profile
sets A1:M31 to include the banner and fits the form to one page. Other templates
must be checked for drawings anchored outside their print area. The bundled
blank-template PDF uses these same explicit profile print settings; the original
Excel template itself remains unchanged. To regenerate it, use:

```sh
python scripts/coe_tos.py preview-template --out blank-template.pdf
```

Review the native export page by page. In particular check:

- Banner, logo and control box are visible.
- Program/course fields and item references are unclipped.
- Signature/approval areas and footnotes are present.
- Percent formatting and conditionally highlighted bands are correct.
- Multi-page content is in order and tables do not leave isolated rows.

Static formula/cache fidelity does not prove visual layout. `--render` writes
page PNGs through optional PyMuPDF and leaves `visual_review` pending. The agent
must inspect those PNGs and add its review evidence to the report. On hosts
without visual tools, disclose that review is still pending.

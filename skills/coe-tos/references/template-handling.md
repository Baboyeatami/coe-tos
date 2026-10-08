# Template preservation and native PDF export

## Strategy

An XLSX is an OOXML ZIP package. `workbooks.py` reads the supplied workbook,
changes cell values and supported formula caches in worksheet XML, and updates
only the profile's print area, fit-to-page properties and calculation flags.
It writes a staged package and checks fidelity before replacing the output.
Other non-edited parts retain their original bytes, including styles, grouped
drawings, images, VML comments, themes and document metadata. Added supplementary
sheets require declared workbook/content-type/relationship registrations.

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

Default builds append generated worksheets (`Assessment mapping`, `Topic
allocation`, `Allocation ledger`, `Notes`) to the delivered workbook. They are
written straight into the OOXML package: a new worksheet part plus the matching
entries in `xl/workbook.xml`, `xl/_rels/workbook.xml.rels` and
`[Content_Types].xml`. No library save round trip occurs. The form retains its
original structure/styles outside declared value/cache/print edits; styles.xml,
drawings and media remain byte-identical. --form-only omits supplementary tabs.

The new sheets use inline strings and no formulas, so nothing needs
recalculating. Percentages are written as preformatted text because the
template's style table is deliberately left untouched.

The fidelity gate treats these additions as declared: `verify()` requires the
set of added parts to match the declared additions exactly, removes only the declared
sheet entries and relationships before comparing, and still fails on any other
package, worksheet or workbook change. An undeclared sheet is refused.

They are visible by default for Excel-only delivery. --pdf hides them to keep
native printing to the form; --mapping-visible keeps them visible in PDF mode.
--hide-mapping explicitly hides them in Excel-only mode. Checks are embedded on
Notes; fidelity.json is written only when --diagnostics is requested.

Once a sheet is appended, the workbook is no longer byte-identical to the
institutional original. The standard Excel deliverable includes these tabs by
default; use --form-only when a strict form-only delivery is required.

## Test Item No. cells

The institutional row height is fixed, so a topic that maps to several criteria
can overflow its `Test Item No.` cell. Set `refs_max_chars` in the profile to the
usable width (the bundled CJC profile uses 16). References are then abbreviated as
`Q1-c1b +2`: as many as fit, plus the number omitted. No information is lost,
because the full list appears in the allocation ledger, in the generated
`Topic allocation` sheet's Refs column. Omit `refs_max_chars`
to print every reference in full.

## Formulas

The deterministic evaluator supports cell/range references, cross-sheet references,
arithmetic, comparisons, IF, SUM, AND and OR, covering the bundled template.
Other formula caches are cleared and listed on Notes, with native
recalculation enabled. The formula XML remains intact. Formula evaluation has
no Python `eval` or execution of worksheet content.

## Print and visual checks

`excel-windows` uses desktop Excel COM automation through optional pywin32, in a
supervised worker with a 180-second timeout. It opens a temporary copy read-only,
disables link updates, recalculates, and exports with the existing print areas and
sheet visibility. The worker owns a new Office instance; timeout cleanup verifies
both its process ID and creation time, and leaves unrelated user instances alone.

`excel-mac` drives an installed Microsoft Excel through AppleScript on a temporary
copy and closes only the workbook it opened. macOS may request automation
permission. `libreoffice` uses an isolated temporary user profile and native Calc
PDF export. No renderer saves back over the verified workbook. `auto` prefers
Excel on Windows/macOS and otherwise uses LibreOffice; a selected engine's failure
is reported without silently trying another engine.

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

Static formula/cache fidelity does not prove visual layout. Optional `--pdf --render` writes
page PNGs through optional PyMuPDF and leaves `visual_review` pending. The agent
must inspect those PNGs and record actual findings on Notes or in chat. No default
Excel-only build claims that it performed a visual review. Use temporary previews
only for new templates/layout concerns, then clean them after inspection.

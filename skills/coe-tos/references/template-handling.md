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

# Version 1.0.0 verification

Local verification performed on macOS with Python 3.12 in a fresh virtual
environment using the repository requirements. Optional PyMuPDF was installed
for page rendering; Microsoft Excel provided native PDF export.

## Automated checks

`python -m unittest discover -s tests -v`: **15 tests passed**.

Covered behaviors:

- Correct 16/24/60 cognitive totals and 50/50 question totals.
- Changed published subcriterion scores rejected even when criterion/exam totals match.
- Duplicate allocations, absent evidence, invalid/nonfinite/negative scores rejected.
- Capacity overflow rejected; cognitive-band conflicts retained as warnings.
- DOCX tables, UTF-16 text and XLSX cell extraction.
- Blank/scanned PDF pages flagged for visual inspection/OCR.
- Failed binary extraction recorded with a nonzero CLI status.
- All supported formula caches agree with an independent recomputation from the workbook.
- Original template bytes remain unchanged; drawing/media/style/relationship parts retained.
- Formula overwrite and invalid validation-list values refused.
- Altered template cell styles detected.
- A second sheet/profile layout builds correctly, including numeric summary values.
- Blank template preview only changes explicit print/calculation settings.
- Skill ZIP integrity, full-folder installation and standalone execution outside the repository.

Agent Skills frontmatter was additionally parsed as YAML and checked for the
required lowercase name, description/compatibility limits and linked resources.

## Native output

The synthetic `examples/draft.json` produced:

- `TOS.xlsx`: 19 original package parts retained, 58 original formulas retained,
  no unsupported formula caches, original drawing/media parts byte-identical.
- `TOS.pdf`: **one landscape page**, 936 × 612 points, searchable text.
- Review, mapping and reproducibility artifacts.

The exported page image was inspected. The CJC banner, logo and document-control
box are visible. All ten topics, short program codes, question references,
signature fields and footer fit. Totals show **16 / 16.0%, 24 / 24.0%,
60 / 60.0%, 100 / 100%**, with no cognitive-band warning highlights.

This establishes the tested output's arithmetic and visible template handling.
It does not certify institutional approval or model-independent educational
judgment. The example is a proposed assessment with explicitly scored recall,
not a reclassification of an unchanged existing exam.

The bundled blank-template PDF was exported using the CJC profile's A1:M31
one-page print settings on a temporary copy, retaining the original XLSX.

## Platform scope

GitHub Actions is configured to run the portable suite on Ubuntu, Windows and
macOS with Python 3.10 and 3.12. Native Excel-mac PDF export was exercised locally.
LibreOffice, OCRmyPDF and Tesseract adapters require those applications and
were not exercised during the initial local verification. Installation paths
for OpenCode, Codex and Claude Code follow their published documentation;
other hosts use the generic/manual instruction-loading route.

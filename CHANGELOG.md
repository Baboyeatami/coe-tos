# Changelog

## 1.4.0

- Default builds deliver only TOS.xlsx, with visible assessment mapping, topic
  allocation, evidence-bearing allocation ledger and Notes tabs.
- PDF/native rendering is opt-in with --pdf; external reports require --diagnostics.
  --xlsx-only and --with-mapping remain compatibility aliases; --form-only preserves
  the old form-only behavior.
- Source names/locators, metadata, assumptions, band results and preservation/cache
  checks are embedded in the workbook. Output layout inspection is not falsely
  claimed by a structural-only build.
- Reuse template bytes and an allocation index; avoid repeated package reads and
  repeated reference scans. Supported formula caches are read back before publishing.
- Reuse existing template header/wrap styles on supplementary tabs without changing
  styles.xml. Preserve decimal values and use valid Excel column names beyond AZ.
- Updated skill workflow to use temporary work files, settle scoring before building,
  and distinguish truncated previews from genuine extraction gaps.

## 1.3.1

- Fixed malformed cell references when a supplementary-sheet row was supplied as a
  bare string; such rows produced a workbook that openpyxl and Excel could not read.
  Rows are now normalised, and non-sequence rows are refused with a clear error.
- Added `refs_max_chars` to the profile. A topic mapping to several criteria can now
  overflow the fixed row height, so Test Item No. cells abbreviate to `Q1-c1b +2`
  instead of wrapping and clipping. The full list appears in the allocation ledger,
  the Topic allocation sheet's new Refs column, and `review.md`.

## 1.3.0

- Added optional supplementary worksheets to the delivered workbook:
  assessment mapping, topic allocation, allocation ledger and notes.
  `build --with-mapping` generates them from the validated ledger; they are
  hidden by default so the exported form PDF stays one page, and
  `--mapping-visible` prints them.
- Sheets are written into the OOXML package directly, so the template's own
  worksheet, styles, drawings and media stay byte-identical.
- The fidelity gate now accounts for declared additions and still refuses any
  undeclared sheet, part or workbook change.
- `fidelity.json` reports `added_sheets` and `added_parts`, and `review.md`
  states when supplementary worksheets were added.

## 1.2.0

- Added `ingest-text` so text pasted in chat joins the workflow with the same
  `line N` locators as a file, including `--append` for multiple pastes.
- Documented handling of attachments the model cannot open directly: locate the
  file, extract it, and keep the source traceable.
- Added input-provenance metadata (`input_mode`, `sources`, `totals_confirmed_by`)
  and a warning when pasted input cannot be checked against page images.
- `review.md` now generates an assessment mapping table (published versus ledger
  points per criterion and published subcriterion) and a topic allocation table,
  both derived from the validated ledger.
- Band results in `review.md` show the profile's actual band range.
- Added regression tests for chat ingestion and the generated report sections.

## 1.1.0

- Added render-only PDF previews with content-keyed, hash-verified PNG caching.
- Added adjustable preview scale, backend preflight, and export/render timings.
- Hardened Excel automation with uniquely named temporary workbooks, bounded
  workbook polling, workbook-level native PDF export, and actionable errors.
- Warned when a failed export may leave an older PDF in the output directory.
- Clarified evidence-led cognitive classification and recovery guidance.
- Added regression tests for render caching and export failure handling.

## 1.0.0

- Portable Agent Skills instructions for evidence-led TOS drafting.
- PDF/DOCX/XLSX/text extraction and optional OCR/legacy conversion adapters.
- Question, criterion and published subcriterion scoring validation.
- Direct OOXML value/cache edits with template/drawing/media fidelity checks.
- CJC Excel template/profile and native blank-form PDF reference.
- Native Excel-mac/LibreOffice PDF export and optional page rendering.
- Harness-neutral installation, reproducible examples and cross-platform tests.

# CoE-TOS

Create an evidence-led College of Engineering Table of Specifications from exams,
rubrics, syllabi, supplied documents or text pasted in chat.

By **Engr. Jamie Eduardo Rosal, MSCpE**. Current skill version: **1.4.0**.

**The default deliverable is one file: `TOS.xlsx`.** Assessment mapping, topic
allocation, evidence, rationales, metadata, assumptions and checks are inside it.
Normal builds do not launch Office or create PDFs, PNGs, Markdown, CSV or JSON reports.

![Synthetic example rendered through Microsoft Excel](docs/example-tos.png)

The illustration is the bundled proposed assessment. Educational classification
and institutional approval remain separate from arithmetic checks.

## What it does

- Maps an existing examination or drafts a proposed assessment blueprint.
- Preserves every published question, criterion and subcriterion score.
- Records source names, locators and cognitive rationales in the workbook.
- Populates the supplied institutional form or the bundled Cor Jesu College form.
- Preserves template formulas, styles, drawings, media, validation and protection
  through direct OOXML edits rather than a spreadsheet-library save round trip.
- Checks totals, supported formula caches and package fidelity before publication.
- Reports missing information and cognitive-band conflicts honestly.

The model reads the sources and drafts the ledger; deterministic Python tools
validate it and build the workbook. The skill has no model/provider API dependency.
It does not force an existing examination into cognitive bands by relabelling tasks.

## Quick start

```sh
git clone https://github.com/Baboyeatami/coe-tos.git
```

From the repository root, install the complete skill folder:

| Host | Command | User installation |
|:--|:--|:--|
| OpenCode | `python scripts/install.py --harness opencode` | `~/.config/opencode/skills/coe-tos` |
| Codex | `python scripts/install.py --harness codex` | `~/.agents/skills/coe-tos` |
| Claude Code | `python scripts/install.py --harness claude` | `~/.claude/skills/coe-tos` |
| Other host | `python scripts/install.py --harness generic --destination /path/to/skills/coe-tos` | Specified folder |

Project-scoped installations support `--scope project --project /path/to/project`.
Use `--force` to upgrade an inspected installation. Restart OpenCode after an
upgrade. Keep scripts, references, assets and licensing with SKILL.md; copying
only the instruction file is insufficient. Installable archives can be found on
the [Releases page](https://github.com/Baboyeatami/coe-tos/releases).

Python 3.10 or later:

```sh
python -m venv .venv
# macOS/Linux: source .venv/bin/activate
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r skills/coe-tos/scripts/requirements.txt
```

Tell the agent to use that environment's Python interpreter. Excel-only delivery
does not require Microsoft Excel, LibreOffice or PyMuPDF. Optional tools are:

- Microsoft Excel on macOS or LibreOffice for requested native PDF export.
- PyMuPDF for source/layout PNG previews: `python -m pip install pymupdf`.
- OCRmyPDF/Tesseract for scanned inputs, or the host's vision tools.

### Ask the agent

> Use CoE-TOS to map this exam and rubric to the CJC form. Deliver only TOS.xlsx,
> including assessment mapping, topic allocation and source evidence in visible tabs.

You can attach PDF/Word/Excel files, give local paths, or paste the exam and scoring
text into the conversation. When the model cannot view an attachment directly, it
can still resolve the file and use local extraction. Text-only chat with no file
tools can produce a draft, but cannot create an actual workbook.

For Codex use `$coe-tos`; for Claude Code use `/coe-tos`. Other hosts can read
[`skills/coe-tos/SKILL.md`](skills/coe-tos/SKILL.md) directly. See the discovery docs
for [OpenCode](https://opencode.ai/docs/skills/),
[Codex](https://developers.openai.com/codex/build-skills), and
[Claude Code](https://code.claude.com/docs/en/skills).

## Excel-only workflow

Keep working files outside the final delivery folder:

```sh
# Use extract for files, or ingest-text for pasted content; these are alternatives.
python skills/coe-tos/scripts/coe_tos.py extract exam.pdf rubric.docx --out work/sources.json
python skills/coe-tos/scripts/coe_tos.py ingest-text --name exam.txt --stdin --out work/sources.json
python skills/coe-tos/scripts/coe_tos.py inspect-template template.xlsx --out work/template-inspection.json
# The agent drafts work/draft.json after reading the sources and inspecting the form.
python skills/coe-tos/scripts/coe_tos.py build --draft work/draft.json --out output
```

For a different form add `--template template.xlsx --profile work/profile.json`.
`build` already validates; a second unchanged `validate` run is unnecessary.
Clean temporary work files after successful verification; never remove original
input documents, institutional templates or unknown user files.

Bundled demonstration:

```sh
python skills/coe-tos/scripts/coe_tos.py build --draft examples/draft.json --out output/demo
```

It proposes two 50-point questions and explicitly weighted recall/application/
analysis tasks, producing 16/24/60. It is not evidence for classifying another exam.

### What is inside TOS.xlsx?

| Visible tab | Contents |
|:--|:--|
| Table of Specifications | Institutional form with item references, points and percentages |
| Assessment mapping | Published versus allocated criterion/subcriterion scores, R/U/T and partition labels |
| Topic allocation | Topic totals, percentages and full references |
| Allocation ledger | Every allocation, rationale, source name and source locator |
| Notes | Metadata, assumptions, proposed changes, cognitive bands, formula-cache and template checks |

The output folder contains only **TOS.xlsx** on a fresh default build. The builder
does not automatically delete existing files in a reused folder. All tabs are
visible for Excel-only delivery; no unhide step is needed.

### Upgrading from 1.3.x

The output defaults changed in 1.4.0:

- A normal `build` delivers the complete Excel workbook; `--xlsx-only` is no longer
  a partial-delivery option.
- Mapping tabs are included and visible without `--with-mapping`.
- Use `--pdf --render` if you explicitly need the old PDF/PNG export workflow.
  `--render` alone is rejected to avoid an unintended Office launch.
- Use `--diagnostics` if you explicitly need the separate Markdown/CSV/JSON files.
- Use `--form-only` when the delivered workbook must contain only the institutional
  form. Use `--hide-mapping` if you want the supplementary tabs hidden.

To upgrade an existing OpenCode installation from an inspected repository checkout:

```sh
git pull --ff-only
python scripts/install.py --harness opencode --force
```

Restart OpenCode to reload the new skill instructions.

### Generation time

Excel-only generation avoids native Office export, which was the main measured
build cost. The builder also reuses template bytes and a shared allocation index,
then verifies the staged workbook's values and supported formula caches before
publication. These checks remain enabled in the faster path.

On the development Mac, three-run medians for the existing 50-point Embedded
Systems and 100-point Engineering Data Analysis drafts were approximately **0.15
seconds per workbook**, including Python startup. These are sample generation
measurements, not guarantees or end-to-end assessment-analysis timings. Source
extraction, drafting and any requested visual inspection are additional work.

## Optional outputs

| Option | Effect |
|:--|:--|
| `--pdf` | Explicitly add a native PDF; requires Office. Mapping tabs are hidden for form-only printing. |
| `--pdf --render` | Also create PNG previews; requires PyMuPDF. |
| `--diagnostics` | Explicitly retain JSON/Markdown/CSV reports, including draft/profile/fidelity/timings. |
| `--form-only` | Omit supplementary tabs for a strict institutional form. |
| `--hide-mapping` | Hide supplementary tabs explicitly. |
| `--mapping-visible` | Keep mapping tabs visible even when exporting a PDF. |
| `--xlsx-only`, `--with-mapping` | Compatibility aliases for the default Excel-only/mapping behavior. |

For a layout concern or new template, export to a temporary location, inspect the
images, then remove those temporary files. A structural check never means a visual
review was performed. Source tables/equations still need visual checks as required.

```sh
python skills/coe-tos/scripts/coe_tos.py export-pdf output/TOS.xlsx --out work/preview.pdf --render
python skills/coe-tos/scripts/coe_tos.py render-pdf work/preview.pdf --scale 2
```

PNG caching verifies PDF content, scale, renderer version and image hashes. Cached
previews still need inspection. Requested exports operate on copies of the workbook.

## Command reference

| Command | Purpose |
|:--|:--|
| `extract` | Local file extraction with locators; `--ocr` for supported scans |
| `ingest-text` | Chat text, with `--stdin` or a file, `--name`, and `--append` |
| `inspect-template` | Actual labels, merges, formulas, validations and geometry |
| `validate` | Explicit arithmetic/band report when needed while drafting |
| `build` | Default single-file Excel delivery; evidence and checks are embedded |
| `export-pdf` | Optional native PDF of an existing workbook |
| `render-pdf` | Optional cached PNG previews of an existing PDF |
| `preview-template` | Optional native preview of a blank form |

## Templates and evidence

The bundled CJC form has ten topic rows and three groups: Remembering,
Understanding (Comprehension/Application), and Thinking (Analysis/Synthesis/
Evaluation). Its bands are R 10–20%, U 20–30%, T at least 60%; they are institution
specific. A different template needs an inspected profile and its own bands.

The CJC semester values are exactly ` 1st` (leading space), `2nd`, `Summer`.
Multi-program codes are written in C7 while leaving the validated A7 label intact.
Full program names and metadata provenance remain on Notes.

Scores must be allocated once. Preserve published subcriteria even when the total
already matches. Mark unscored within-criterion splits as proposed for adoption;
do not invent recall points to satisfy a band. Cognitive classifications require
instructor review. Source data, not a truncated screen preview, determines whether
information is missing.

The form's existing formula/style structure is preserved. Other non-edited package
parts, including styles and drawings/media, remain byte-identical. Added worksheets
and their registrations are declared exceptions checked by the fidelity gate.
No default Excel build claims pixel-identical rendering or visual approval.

Supported formula evaluation includes IF, SUM, AND, OR, arithmetic, comparisons
and references. Unsupported caches are cleared, disclosed on Notes and left for
native recalculation rather than guessed. See
[template handling](skills/coe-tos/references/template-handling.md) and
[draft format](skills/coe-tos/references/draft-format.md).

## Tests and packaging

```sh
python -m unittest discover -s tests -v
python scripts/package_skill.py --out dist/coe-tos.zip
```

Tests cover extraction, scoring/subcriteria, template fidelity, formula caches,
chat ingestion, mapping sheets, render caching and single-file Excel delivery.
The Excel-only build is tested without invoking Office. Optional PDF smoke tests
need an installed renderer. GitHub Actions tests Python 3.10/3.12 on Windows,
Linux and macOS; PyMuPDF-specific tests are skipped if the optional package is absent.

## License

Instructions, tools and authored examples: [MIT](LICENSE). Institutional artwork,
logos and names retain their owners' rights; see
[third-party notices](THIRD_PARTY_NOTICES.md). The CJC form does not imply endorsement.

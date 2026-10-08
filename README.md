# CoE-TOS

Create an evidence-led College of Engineering Table of Specifications from exams,
rubrics, syllabi, supplied documents or text pasted in chat.

By **Engr. Jamie Eduardo Rosal, MSCpE**. Current skill version: **1.5.0**.

**Already installed an older version?** Check yours, then follow
[Updating an existing installation](#updating-an-existing-installation).
The [1.0.x migration notes](#updating-from-10x) explain changed output defaults;
[1.5.0 adds Windows desktop Excel and Word support](#updating-from-14x).

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

- Desktop Microsoft Excel on Windows or macOS for requested native PDF export;
  LibreOffice is the fallback. Windows also needs the optional pywin32 dependency.
- Desktop Microsoft Word/Excel on Windows for legacy `.doc`/`.xls` input conversion.
- PyMuPDF for source/layout PNG previews: `python -m pip install pymupdf`.
- OCRmyPDF/Tesseract for scanned inputs, or the host's vision tools.

### Windows with Microsoft Office

From the repository root, use the virtual environment's interpreter directly in
PowerShell (activation is optional):

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r skills/coe-tos/scripts/requirements.txt
.\.venv\Scripts\python.exe -m pip install -r skills/coe-tos/scripts/requirements-windows-office.txt
.\.venv\Scripts\python.exe scripts/install.py --harness opencode --force
```

Tell the agent to use `.venv\Scripts\python.exe`. Desktop Excel/Word must be
installed and registered for automation; browser-only Microsoft 365 does not
provide COM automation. The optional dependency is needed for Office operations,
while default `TOS.xlsx` generation works with the core requirements alone.

PDF `--backend auto` prefers Excel on Windows and macOS; Linux uses LibreOffice.
The explicit choices are `excel-windows`, `excel-mac`, and `libreoffice`:

```powershell
.\.venv\Scripts\python.exe skills/coe-tos/scripts/coe_tos.py build --draft examples/draft.json --out output
.\.venv\Scripts\python.exe skills/coe-tos/scripts/coe_tos.py export-pdf output/TOS.xlsx --out work/preview.pdf --backend excel-windows
```

Legacy `.doc` and `.xls` extraction automatically prefers Windows Word and Excel,
respectively. Use `extract ... --office-backend ms-office` to require that engine,
or `--office-backend libreoffice` to select LibreOffice explicitly. `.odt`, `.ods`
and `.rtf` conversion uses LibreOffice. Modern PDF/DOCX/XLSX inputs use native
Python readers. An export/conversion error is reported without silently switching
engines.

LibreOffice discovery checks `PATH` and standard Windows Program Files locations.
An explicit override is available for custom installations:

```powershell
$env:COE_TOS_SOFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'
```

Windows Office runs in a supervised worker with a 180-second timeout, using copies
of input files and a newly created Office instance. Cleanup checks the recorded
process ID and creation time. If Excel/Word reuses an existing process, the worker
refuses to operate on it. See [installation and troubleshooting](skills/coe-tos/references/installation.md).

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

## Updating an existing installation

Check which version you are running before anything else. The version is the
`version` field in the frontmatter of the installed `SKILL.md`:

```sh
grep -m1 'version:' ~/.config/opencode/skills/coe-tos/SKILL.md   # OpenCode
grep -m1 'version:' ~/.agents/skills/coe-tos/SKILL.md           # Codex
grep -m1 'version:' ~/.claude/skills/coe-tos/SKILL.md           # Claude Code
```

Windows PowerShell, for an OpenCode user installation:

```powershell
Select-String -Path "$HOME\.config\opencode\skills\coe-tos\SKILL.md" -Pattern '^\s*version:'
```

Update older installations to **1.5.0** for the current delivery workflow and
Windows Office support. From your repository checkout:

```sh
git clone https://github.com/Baboyeatami/coe-tos.git   # first time only
cd coe-tos
git pull --ff-only                                     # existing checkout
python scripts/install.py --harness opencode --force    # see the harness table
```

`--force` is required whenever the destination already exists; the installer
refuses to overwrite a folder that has no `SKILL.md`, and refuses a symlinked
destination rather than replacing it. Inspect the target folder first if you
stored local edits there — `--force` deletes it and reinstalls from the repository.
Use `--scope project --project /path/to/project` for a project-scoped install, or
`--harness generic --destination /path/to/skills/coe-tos` for another host.

From a release archive instead of a checkout, unpack the new version over a copy
of the same skill folder, keeping `scripts/`, `references/`, `assets/`, `LICENSE`
and `SKILL.md` together, then restart the host. Replacing only `SKILL.md` leaves
an older script set in place and will fail.

Restart OpenCode, or reload the host, after upgrading. A running session keeps the
old instructions in memory. The core dependency list is unchanged since 1.0.0;
install the optional Windows Office requirements into the interpreter used by
your agent if you want native Excel/Word automation:

```powershell
python -m pip install -r skills/coe-tos/scripts/requirements-windows-office.txt
```

### Updating from 1.0.x

Your existing `draft.json` files, custom `--profile` JSON and institution
templates keep working. The draft schema only gained optional fields (`title`,
question `title`, and the `input_mode` / `sources` / `totals_confirmed_by` /
`totals_confirmed_on` provenance metadata), and `scoring.py` gained checks without
changing any rule. The bundled `examples/draft.json` is byte-identical to 1.0.0
and still builds to 16 / 24 / 60. Existing valid drafts need no schema migration.
The behavioural differences that affect work you already did:

| | 1.0.x behaviour | Current behaviour |
|:--|:--|:--|
| Default `build` output | Workbook, PDF and external reports; PNGs when `--render` was requested | **`TOS.xlsx` only**; evidence and checks are inside it as visible tabs |
| PDF export | Attempted by default; without a renderer it exited nonzero | Opt-in via `--pdf`, so a normal build never launches Office |
| `build … --render` | The documented command | **Rejected** — `--render` now requires `--pdf` |
| `--xlsx-only` | Skipped PDF export on a machine with no renderer | Compatibility alias for the new default |
| Supplementary tabs | Did not exist (1.0.0) | Present and visible without `--with-mapping` |
| Working files | Often written into `output/` | Keep them in a temporary work folder outside the delivery folder |

So, concretely:

- Re-run the 1.0 command **with `--pdf`** if you want the PDF and PNGs back:
  `--render` on its own now exits with an error.
- Re-run with `--diagnostics` if your process or review relied on `review.md`,
  `mapping.csv` and the other external reports. Do not expect them by default.
- Use `--form-only` for a strict institutional form, or `--hide-mapping` to keep the
  supplementary tabs out of sight while retaining them.
- Keep `draft.json`, `sources.json` and `template-inspection.json` in a work folder,
  not in the folder you deliver. The builder does not delete files it did not create,
  so old `review.md` / `fidelity.json` files from a 1.0 run will otherwise sit next to
  the new `TOS.xlsx` and look current.

Two corrections arrive with the upgrade. A topic mapping to several criteria used to
wrap and clip in the fixed-height Test Item No. cell; cells now abbreviate as
`Q1-c1b +2` and the full list stays in the Topic allocation Refs column and the
Allocation ledger. A supplementary-sheet row supplied as a bare string produced a
workbook that Excel and openpyxl could not open; rows are now normalised. Rebuild
affected outputs with clipped references or unreadable 1.3.0 supplementary sheets.

If you previously asked for "the XLSX, PDF and review report", ask for
"only `TOS.xlsx`, with the mapping, evidence and notes in visible tabs" to get the
current default.

### Updating from 1.3.x

The output defaults changed in 1.4.0:

- A normal `build` delivers the complete Excel workbook; `--xlsx-only` is no longer
  a partial-delivery option.
- Mapping tabs are included and visible without `--with-mapping`.
- Use `--pdf --render` if you explicitly need the old PDF/PNG export workflow.
  `--render` alone is rejected to avoid an unintended Office launch.
- Use `--diagnostics` if you explicitly need the separate Markdown/CSV/JSON files.
- Use `--form-only` when the delivered workbook must contain only the institutional
  form. Use `--hide-mapping` if you want the supplementary tabs hidden.

Install with the same command as above (`git pull --ff-only`, then
`python scripts/install.py --harness opencode --force`) and restart OpenCode.

### Updating from 1.4.x

Version 1.5.0 adds `--backend excel-windows` and Windows native `.doc`/`.xls`
conversion. Install `requirements-windows-office.txt` in your existing environment,
then reinstall the entire skill folder with `--force` and restart the host. The
one-file Excel default, draft format and assessment mapping are retained.

## Generation time

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
The Excel-only build is tested without invoking Office. Cross-platform tests cover
backend selection, COM method contracts, process ownership, failures, and a real
subprocess watchdog using a non-Office fixture. GitHub Actions tests Python
3.10/3.12 on Windows, Linux and macOS; PyMuPDF-specific tests are skipped if the
optional package is absent.

Real Windows Office tests are **opt-in**, and require desktop Excel and Word:

```powershell
$env:COE_TOS_RUN_OFFICE_TESTS = '1'
python -m unittest discover -s tests -p 'test_windows_office.py' -v
Remove-Item Env:COE_TOS_RUN_OFFICE_TESTS
```

These check form-only and mapping-visible PDF export, preservation of an unrelated
open workbook, and `.doc`/`.xls` conversion. They remain skipped in ordinary CI;
mocked tests do not establish real Windows Office compatibility. Inspect native
PDF pages for layout, which can vary with fonts and printer settings.

## License

Instructions, tools and authored examples: [MIT](LICENSE). Institutional artwork,
logos and names retain their owners' rights; see
[third-party notices](THIRD_PARTY_NOTICES.md). The CJC form does not imply endorsement.

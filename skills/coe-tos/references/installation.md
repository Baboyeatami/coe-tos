# Installation and capabilities

The portable unit is the entire `coe-tos/` directory containing SKILL.md,
scripts, references and assets. Copying only SKILL.md loses templates/tools.

From a repository checkout:

```sh
python scripts/install.py --harness codex
python scripts/install.py --harness opencode
python scripts/install.py --harness claude
```

Use just the command for your host. Global locations:

- Codex: `~/.agents/skills/coe-tos`
- OpenCode: `~/.config/opencode/skills/coe-tos`
- Claude Code: `~/.claude/skills/coe-tos`

Project installation: add `--scope project --project /path/to/project`.
Other hosts: `--harness generic --destination /path/to/skills/coe-tos`.
Inspect existing installations before using `--force` to replace one.

Dependencies, using the actual installation path:

```sh
python -m venv .venv
# macOS/Linux: source .venv/bin/activate
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r /path/to/coe-tos/scripts/requirements.txt
```

Python 3.10+ is required. `python3` on Unix and `py` on Windows may be the
correct interpreter command. The host must use that environment/interpreter
when it runs scripts. Optional: `pip install pymupdf` for page PNGs, OCRmyPDF
plus its documented system dependencies for scanned PDFs, Tesseract for images.

## Microsoft Office on Windows

Install desktop Excel for native PDF export and desktop Word for legacy `.doc`
conversion. Browser-only Microsoft 365 does not provide local COM automation.
From the repository root, PowerShell can use the environment directly:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r skills/coe-tos/scripts/requirements.txt
.\.venv\Scripts\python.exe -m pip install -r skills/coe-tos/scripts/requirements-windows-office.txt
.\.venv\Scripts\python.exe scripts/install.py --harness opencode --force
```

For an already-installed skill, use its actual path to each requirements file.
The agent must run this environment's interpreter. Core XLSX delivery needs no
Office installation or pywin32.

| Operation | Windows | macOS | Linux |
|:--|:--|:--|:--|
| Default `TOS.xlsx` build | Native Python tools | Native Python tools | Native Python tools |
| Native PDF | Desktop Excel + pywin32; LibreOffice fallback | Desktop Excel via AppleScript; LibreOffice fallback | LibreOffice |
| Legacy `.doc` / `.xls` extraction | Word / Excel + pywin32; LibreOffice fallback | LibreOffice or manually save as DOCX/XLSX | LibreOffice |
| `.odt` / `.ods` / `.rtf` extraction | LibreOffice | LibreOffice | LibreOffice |

PDF backend choices are `auto`, `excel-windows`, `excel-mac`, and `libreoffice`,
accepted by `build`, `export-pdf`, and `preview-template`. `auto` prefers an
available native Excel backend; explicit choices fail if unavailable. Excel is
detected by COM registration on Windows, without starting it during preflight.
An operation that fails after selecting an engine is reported without trying
another engine silently.

```powershell
.\.venv\Scripts\python.exe skills/coe-tos/scripts/coe_tos.py export-pdf output/TOS.xlsx --out work/preview.pdf --backend excel-windows
.\.venv\Scripts\python.exe skills/coe-tos/scripts/coe_tos.py extract exam.doc rubric.xls --office-backend ms-office --out work/sources.json
```

`extract --office-backend auto` prefers Word for `.doc` and Excel for `.xls` on
Windows. `--office-backend libreoffice` explicitly selects LibreOffice. Source
warnings name the converter and disclose that locators refer to converted content.

## LibreOffice discovery

The scripts check `COE_TOS_SOFFICE`, then `soffice`/`libreoffice` on PATH. Windows
also checks `ProgramW6432`, `ProgramFiles`, and `ProgramFiles(x86)` under
`LibreOffice\program\soffice.exe`; macOS checks the standard application bundle.
For a custom installation, point the override at the executable:

```powershell
$env:COE_TOS_SOFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'
```

An invalid explicit override produces an error, rather than being ignored.

## Delivery and previews

Version 1.4 defaults to complete Excel-only delivery:

```sh
python scripts/coe_tos.py build --draft work/draft.json --out output
```

This writes only TOS.xlsx with visible mapping, ledger/evidence and Notes tabs.
Office/PyMuPDF are not loaded by the default build. `--pdf` explicitly adds native
PDF export; `--pdf --render` creates previews. `--diagnostics` explicitly keeps the
external reports. `--xlsx-only` is a compatibility alias, not a partial delivery.

Rendering dependencies are checked before opening Office. For PNG-only previews
of an existing PDF, use `render-pdf` instead of running `build` again:

```sh
python scripts/coe_tos.py render-pdf output/TOS.pdf --scale 1
```

The default scale is 1.5; use 2 for closer inspection. This command does not
require Excel or LibreOffice. It reuses only content/version/scale-matched page
images whose hashes still match. A preview cache hit is not visual approval.

Text supplied directly in chat is ingested with `ingest-text`, which needs no
extra dependency and records `line N` locators:

```sh
python scripts/coe_tos.py ingest-text --name exam-rubric.txt --stdin --out sources.json
```

It adds a warning that pasted text cannot be visually verified, and
`draft.json` metadata `input_mode: chat-pasted` repeats that caution in
the workbook's Notes tab (and `review.md` only with `--diagnostics`).

Run rendering regression tests from the repository root:

```sh
python -m unittest discover -s tests -p 'test_export_pdf.py' -v
```

Render-specific tests use PyMuPDF when installed and are skipped otherwise. They
cover changed PDFs, corrupt cached PNGs, scale changes, dependency failures before
Office launch, and Excel error reporting/ownership. The full suite also covers
Windows backend selection, COM method contracts, PID/creation-time cleanup, and
a subprocess timeout fixture without starting Office.

Real Windows Office tests require an explicit opt-in on a desktop with Excel and
Word installed. They check form-only/mapping-visible PDFs, an unrelated workbook
remaining open, and legacy Word/Excel conversion:

```powershell
$env:COE_TOS_RUN_OFFICE_TESTS = '1'
python -m unittest discover -s tests -p 'test_windows_office.py' -v
Remove-Item Env:COE_TOS_RUN_OFFICE_TESTS
```

These native tests are skipped by default, including hosted CI. Passing mocked
tests is not a claim that a real Windows Office export or visual review occurred.

## Windows troubleshooting

- **pywin32 missing:** install the optional requirements with the same Python
  interpreter used by the agent; installing into another Python does not help.
- **Excel/Word not registered:** open the desktop application manually once to
  finish first-run/licensing setup. Repair the Office installation if its COM
  registration remains unavailable.
- **Timeout:** the parent bounds each native operation at 180 seconds. It cleans
  up only the new process identified by PID and creation time. If the worker was
  blocked before recording ownership, it reports that cleanup is unverified;
  inspect Office manually. Existing user instances are never globally terminated.
- **Existing process reused:** the worker refuses to automate/quit that instance.
  Resolve the Office registration/startup issue or explicitly select LibreOffice.
- **Different layout:** fonts, printer metrics and native rendering can differ
  between platforms. Inspect the PDF pages; equal scores do not prove equal layout.

Restart OpenCode after installing. Codex discovers skills automatically;
restart if missing. Claude Code supports reload/restart when discovery fails.
Invoke as `$coe-tos` in Codex, `/coe-tos` in Claude Code, or ask OpenCode to
use the coe-tos skill. Any other agent can read SKILL.md and follow it directly.

Text-only chat: paste/upload the instructions and source text to obtain a
draft/review. A file-generation sandbox is needed to create XLSX/PDF; instructions
alone do not give a model filesystem or Office tools. Model-agnostic does not
mean every model has identical document understanding or tool capabilities.

No model API key is required by these scripts. Sources are processed locally;
an agent's own provider/data handling is governed by its host configuration.

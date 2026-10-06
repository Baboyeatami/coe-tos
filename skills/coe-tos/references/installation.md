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

For PDF export install Microsoft Excel on macOS or LibreOffice on macOS,
Linux or Windows. Put `soffice` on PATH (Windows normally installs it under
`C:\Program Files\LibreOffice\program`). XLSX generation does not require Office.

Rendering dependencies are checked before opening Office. For PNG-only previews
of an existing PDF, use `render-pdf` instead of running `build` again:

```sh
python scripts/coe_tos.py render-pdf output/TOS.pdf --scale 1
```

The default scale is 1.5; use 2 for closer inspection. This command does not
require Excel or LibreOffice. It reuses only content/version/scale-matched page
images whose hashes still match. A preview cache hit is not visual approval.

Run rendering regression tests from the repository root:

```sh
python -m unittest discover -s tests -p 'test_export_pdf.py' -v
```

Render-specific tests use PyMuPDF when installed and are skipped otherwise. They
cover changed PDFs, corrupt cached PNGs, scale changes, dependency failures before
Office launch, and Excel error reporting/ownership. Native export still needs a
smoke test on a host with Office installed.

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

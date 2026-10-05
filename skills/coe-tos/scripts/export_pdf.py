"""Native spreadsheet PDF export; never substitute a hand-drawn imitation."""
from pathlib import Path
import json
import platform
import shutil
import subprocess
import tempfile

from pypdf import PdfReader
from documents import convert_office, office_binary


def export(source, output, backend="auto", render=False):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.suffix.lower() != ".pdf":
        raise ValueError("PDF output must be a different file with a .pdf extension")
    output.parent.mkdir(parents=True, exist_ok=True)
    excel = platform.system() == "Darwin" and Path("/Applications/Microsoft Excel.app").exists()
    if backend == "auto":
        backend = "excel-mac" if excel else "libreoffice"
    if backend == "excel-mac" and not excel:
        raise ValueError("excel-mac needs Microsoft Excel on macOS")
    if backend == "libreoffice" and not office_binary():
        raise ValueError("PDF export needs Excel on macOS or LibreOffice. Install a backend or explicitly request --xlsx-only.")
    with tempfile.TemporaryDirectory(prefix="coe-tos-export-") as directory:
        temporary = Path(directory)
        # Rendering engines operate on a copy, never on the verified XLSX artifact.
        copied = temporary / "form.xlsx"
        shutil.copy2(source, copied)
        generated = temporary / "form.pdf"
        if backend == "excel-mac":
            script = f'''
tell application "Microsoft Excel"
    open workbook workbook file name {json.dumps(str(copied))}
    set targetWorkbook to active workbook
    try
        save as (sheet 1 of targetWorkbook) filename {json.dumps(str(generated))} file format PDF file format
    on error messageText number errorNumber
        close targetWorkbook saving no
        error messageText number errorNumber
    end try
    close targetWorkbook saving no
end tell
'''
            subprocess.run(["osascript", "-"], input=script, text=True,
                           capture_output=True, check=True, timeout=180)
        elif backend == "libreoffice":
            generated = convert_office(copied, "pdf:calc_pdf_Export", temporary)
        else:
            raise ValueError(f"Unknown PDF backend: {backend}")
        if not generated.is_file():
            raise ValueError("Spreadsheet renderer did not produce a PDF")
        reader = PdfReader(generated)
        if not reader.pages:
            raise ValueError("Exported PDF is empty")
        page_text = [page.extract_text() or "" for page in reader.pages]
        if not any(text.strip() for text in page_text):
            raise ValueError("Exported PDF has no searchable text; inspect the renderer output")
        pdf_info = {"backend": backend, "pages": len(reader.pages),
                    "page_sizes_points": [[float(p.mediabox.width), float(p.mediabox.height)] for p in reader.pages],
                    "searchable_text": True, "visual_review": "pending", "renders": []}
        shutil.copy2(generated, output)
    if render:
        try:
            import pymupdf
        except ImportError as error:
            raise ValueError("PDF was exported, but --render requires PyMuPDF: pip install pymupdf") from error
        with pymupdf.open(output) as document:
            render_dir = output.parent / (output.stem + "-pages")
            render_dir.mkdir(exist_ok=True)
            for index, page in enumerate(document, 1):
                target = render_dir / f"page-{index:03}.png"
                page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5)).save(target)
                pdf_info["renders"].append(str(target))
    return pdf_info

"""Native spreadsheet PDF export; never substitute a hand-drawn imitation."""
from pathlib import Path
import hashlib
import json
import math
import platform
import shutil
import subprocess
import tempfile
import time
import uuid

from pypdf import PdfReader
from office import PDF_BACKENDS, windows_office_error


def renderer():
    try:
        import pymupdf
        return pymupdf
    except ImportError as error:
        raise ValueError("Page rendering requires PyMuPDF: python -m pip install pymupdf") from error


def render_pdf(source, scale=1.5):
    """Render an existing PDF without Office; reuse only verified, content-keyed PNGs."""
    started = time.perf_counter()
    if not math.isfinite(scale) or not 0.5 <= scale <= 4:
        raise ValueError("Render scale must be between 0.5 and 4")
    fitz = renderer()
    source = Path(source).resolve()
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    version = str(fitz.VersionBind)
    key = hashlib.sha256(f"{digest}:{scale}:{version}".encode()).hexdigest()[:16]
    directory = source.parent / (source.stem + "-pages")
    directory.mkdir(exist_ok=True)
    manifest = directory / f"render-{key}.json"
    try:
        cached = json.loads(manifest.read_text(encoding="utf-8"))
        if (cached["pdf_sha256"] == digest and cached["scale"] == scale
                and cached["renderer_version"] == version and cached["renders"]
                and len(cached["renders"]) == len(cached["image_sha256"])
                and all(Path(path).parent == directory
                        and hashlib.sha256(Path(path).read_bytes()).hexdigest() == image_hash
                        for path, image_hash in zip(cached["renders"], cached["image_sha256"]))):
            return {**cached, "cache_hit": True, "visual_review": "pending",
                    "render_seconds": round(time.perf_counter() - started, 4)}
    except (OSError, ValueError, KeyError, TypeError):
        pass

    renders, hashes = [], []
    with fitz.open(source) as document:
        if not len(document):
            raise ValueError("Cannot render an empty PDF")
        for index, page in enumerate(document, 1):
            target = directory / f"{key}-page-{index:03}.png"
            page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(target)
            renders.append(str(target))
            hashes.append(hashlib.sha256(target.read_bytes()).hexdigest())
    info = {"pdf_sha256": digest, "scale": scale, "renderer_version": version,
            "renders": renders, "image_sha256": hashes, "pages": len(renders),
            "cache_hit": False, "visual_review": "pending",
            "render_seconds": round(time.perf_counter() - started, 4)}
    manifest.write_text(json.dumps(info, indent=2) + "\n", encoding="utf-8")
    return info


def preflight(backend="auto", render=False, scale=1.5):
    """Fail before opening Office when a requested backend/dependency is absent."""
    from documents import office_binary
    if not math.isfinite(scale) or not 0.5 <= scale <= 4:
        raise ValueError("Render scale must be between 0.5 and 4")
    if render:
        renderer()
    excel = platform.system() == "Darwin" and Path("/Applications/Microsoft Excel.app").exists()
    if backend == "auto":
        backend = ("excel-windows" if not windows_office_error("Excel") else
                   "excel-mac" if excel else "libreoffice")
    if backend not in PDF_BACKENDS[1:]:
        raise ValueError(f"Unknown PDF backend: {backend}")
    if backend == "excel-windows":
        error = windows_office_error("Excel")
        if error:
            raise ValueError(error)
    if backend == "excel-mac" and not excel:
        raise ValueError("excel-mac needs Microsoft Excel on macOS")
    if backend == "libreoffice" and not office_binary():
        details = windows_office_error("Excel") if platform.system() == "Windows" else ""
        raise ValueError("PDF export needs desktop Excel on Windows (with pywin32) or macOS, "
                         "or LibreOffice. Omit --pdf for Excel-only delivery. " + details)
    return backend


def export(source, output, backend="auto", render=False, scale=1.5):
    started = time.perf_counter()
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.suffix.lower() != ".pdf":
        raise ValueError("PDF output must be a different file with a .pdf extension")
    backend = preflight(backend, render, scale)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="coe-tos-export-") as directory:
        temporary = Path(directory)
        # Rendering engines operate on a copy, never on the verified XLSX artifact.
        copied = temporary / f"tos-{uuid.uuid4().hex}.xlsx"
        shutil.copy2(source, copied)
        generated = temporary / "form.pdf"
        if backend == "excel-windows":
            from windows_office import run_office
            run_office("pdf", copied, generated)
        elif backend == "excel-mac":
            script = f'''
tell application "Microsoft Excel"
    open (POSIX file {json.dumps(str(copied))})
    repeat with attempt from 1 to 100
        if exists workbook {json.dumps(copied.name)} then exit repeat
        delay 0.1
    end repeat
    if not (exists workbook {json.dumps(copied.name)}) then error "Timed out opening temporary TOS workbook"
    set targetWorkbook to workbook {json.dumps(copied.name)}
    try
        calculate (sheet 1 of targetWorkbook)
        save workbook as targetWorkbook filename {json.dumps(str(generated))} file format PDF file format
    on error messageText number errorNumber
        try
            close targetWorkbook saving no
        end try
        error messageText number errorNumber
    end try
    close targetWorkbook saving no
end tell
'''
            try:
                subprocess.run(["osascript", "-"], input=script, text=True,
                               capture_output=True, check=True, timeout=180)
            except subprocess.CalledProcessError as error:
                raise ValueError(f"Excel PDF export failed: {error.stderr.strip()}") from error
            except subprocess.TimeoutExpired as error:
                raise ValueError("Excel PDF export timed out; inspect the temporary TOS workbook in Excel") from error
        elif backend == "libreoffice":
            from documents import convert_office
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
    pdf_info["export_seconds"] = round(time.perf_counter() - started, 4)
    pdf_info["pdf_sha256"] = hashlib.sha256(output.read_bytes()).hexdigest()
    if render:
        rendered = render_pdf(output, scale)
        pdf_info["renders"] = rendered["renders"]
        pdf_info["render_seconds"] = rendered["render_seconds"]
        pdf_info["render_cache_hit"] = rendered["cache_hit"]
        pdf_info["render_scale"] = scale
    return pdf_info

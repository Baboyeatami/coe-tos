"""Office discovery without launching an application or requiring Office imports."""
import importlib
import os
from pathlib import Path
import platform
import shutil


PDF_BACKENDS = ("auto", "excel-windows", "excel-mac", "libreoffice")
INPUT_BACKENDS = ("auto", "ms-office", "libreoffice")


def office_binary():
    """Find LibreOffice only: its callers pass soffice command-line arguments."""
    override = os.environ.get("COE_TOS_SOFFICE")
    if override:
        path = Path(os.path.expandvars(override)).expanduser().resolve()
        if not path.is_file():
            raise ValueError(f"COE_TOS_SOFFICE does not point to a file: {path}")
        return str(path)
    candidates = [shutil.which("soffice"), shutil.which("libreoffice")]
    if platform.system() == "Windows":
        roots = [os.environ.get(name) for name in
                 ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)")]
        roots += [r"C:\Program Files", r"C:\Program Files (x86)"]
        candidates += [str(Path(root) / "LibreOffice" / "program" / "soffice.exe")
                       for root in roots if root]
    elif platform.system() == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return next((str(path) for path in dict.fromkeys(candidates)
                 if path and Path(path).is_file()), None)


def windows_office_error(application="Excel"):
    """Return an availability error, or None; COM activation happens in the worker."""
    if application not in ("Excel", "Word"):
        raise ValueError(f"Unknown Office application: {application}")
    if platform.system() != "Windows":
        return f"Microsoft {application} COM automation requires Windows"
    try:
        importlib.import_module("pythoncom")
        importlib.import_module("win32com.client")
        registry = importlib.import_module("winreg")
    except (ImportError, OSError) as error:
        return ("Windows Office automation needs pywin32 in this Python environment: "
                "python -m pip install -r scripts/requirements-windows-office.txt "
                f"({error})")
    # Check both registry views for 32-bit Office with 64-bit Python, and vice versa.
    views = (0, registry.KEY_WOW64_64KEY, registry.KEY_WOW64_32KEY)
    for view in views:
        try:
            with registry.OpenKey(registry.HKEY_CLASSES_ROOT,
                                  f"{application}.Application\\CLSID", 0,
                                  registry.KEY_READ | view) as key:
                clsid, _ = registry.QueryValueEx(key, "")
                if clsid:
                    return None
        except OSError:
            continue
    return (f"Desktop Microsoft {application} COM registration was not found. "
            f"Install/repair desktop {application}, or select LibreOffice.")


def input_backend(suffix, backend="auto"):
    if backend not in INPUT_BACKENDS:
        raise ValueError(f"Unknown input conversion backend: {backend}")
    application = {".doc": "Word", ".xls": "Excel"}.get(suffix)
    if backend == "ms-office":
        if not application:
            raise ValueError(f"MS Office input conversion supports .doc and .xls, not {suffix}")
        error = windows_office_error(application)
        if error:
            raise ValueError(error)
        return f"{application.lower()}-windows"
    if backend == "auto" and application and not windows_office_error(application):
        return f"{application.lower()}-windows"
    if not office_binary():
        details = windows_office_error(application) if application else None
        raise ValueError("Legacy input conversion needs Windows desktop Word/Excel with pywin32 "
                         "for .doc/.xls, or LibreOffice. Convert to DOCX/XLSX first."
                         + (f" {details}" if details else ""))
    return "libreoffice"

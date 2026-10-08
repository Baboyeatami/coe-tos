"""Isolated Windows desktop Office operations with a bounded parent watchdog.

The worker owns COM activation; pywin32 is optional and imported lazily. Its PID
record binds cleanup to the newly created Office instance and its creation time,
never to all EXCEL.EXE/WINWORD.EXE processes.
"""
import argparse
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile


APPLICATIONS = {"pdf": "Excel", "xlsx": "Excel", "docx": "Word"}
WORKER = Path(__file__).resolve()


def _write_json(path, value):
    staged = path.with_suffix(".tmp")
    staged.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    staged.replace(path)


def _creation_time(pid):
    import win32api
    import win32process
    # PROCESS_QUERY_LIMITED_INFORMATION, sufficient for GetProcessTimes.
    handle = win32api.OpenProcess(0x1000, False, pid)
    try:
        return str(win32process.GetProcessTimes(handle)["CreationTime"])
    finally:
        handle.Close()


def _quit_application(application, app):
    if application == "Word":
        app.Quit(SaveChanges=0)
    else:
        app.Quit()


def _create_application(application, owner_path):
    import win32com.client
    import win32process

    existing = set(win32process.EnumProcesses())
    app = win32com.client.DispatchEx(f"{application}.Application")
    bootstrap = None
    owned = False
    try:
        if application == "Word":
            # Word exposes HWND on Window, not Application. This document is ours.
            bootstrap = app.Documents.Add()
            hwnd = bootstrap.Windows.Item(1).Hwnd
        else:
            hwnd = app.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if not pid or pid in existing:
            raise ValueError(f"Microsoft {application} reused an existing process; "
                             "refusing to automate or quit a user's Office instance")
        owned = True
        _write_json(owner_path, {"pid": pid, "created": _creation_time(pid),
                                 "application": application})
        if bootstrap is not None:
            bootstrap.Close(SaveChanges=0)
            bootstrap = None
        return app
    except Exception:
        if bootstrap is not None:
            try:
                bootstrap.Close(SaveChanges=0)
            except Exception:
                pass
        if owned:
            try:
                _quit_application(application, app)
            except Exception:
                pass
        raise


def _terminate_owned_process(owner_path):
    """Best-effort cleanup; report uncertainty rather than killing an unverified PID."""
    if not owner_path.is_file():
        return False, "No verified Office PID was recorded; inspect Office for a leftover instance."
    handle = None
    try:
        import win32api
        import win32event
        import win32process
        owner = json.loads(owner_path.read_text(encoding="utf-8"))
        pid = owner["pid"]
        if not isinstance(pid, int) or isinstance(pid, bool) or pid <= 0:
            raise ValueError("Invalid Office PID record")
        try:
            # PROCESS_TERMINATE | SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION.
            handle = win32api.OpenProcess(0x100001 | 0x1000, False, pid)
        except Exception as error:
            if getattr(error, "winerror", None) == 87:  # already exited
                return True, "The owned Office process has exited."
            raise
        created = str(win32process.GetProcessTimes(handle)["CreationTime"])
        if created != owner["created"]:
            return True, "Office PID was reused; the current process was left alone."
        if win32event.WaitForSingleObject(handle, 1000) == 0:
            return True, "The owned Office process has exited."
        win32api.TerminateProcess(handle, 1)
        if win32event.WaitForSingleObject(handle, 5000) != 0:
            return False, "The owned Office process did not exit; inspect Office."
        return True, "The owned Office process was terminated."
    except Exception as error:
        return False, f"Owned Office cleanup could not be verified: {error}"
    finally:
        if handle is not None:
            handle.Close()


def _perform(operation, source, output, owner_path):
    import pythoncom

    application = APPLICATIONS[operation]
    app = document = None
    update_links = None
    pythoncom.CoInitialize()
    try:
        app = _create_application(application, owner_path)
        app.Visible = False
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        if application == "Excel":
            app.DisplayAlerts = False
            app.EnableEvents = False
            document = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True,
                                          IgnoreReadOnlyRecommended=True, AddToMru=False,
                                          Notify=False, Password="", WriteResPassword="")
            if operation == "pdf":
                app.CalculateFull()
                # Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas,
                # From, To, OpenAfterPublish. Use positional typelib arguments.
                document.ExportAsFixedFormat(0, str(output), 0, True, False,
                                             pythoncom.Missing, pythoncom.Missing, False)
            else:
                document.CheckCompatibility = False
                document.SaveAs(str(output), FileFormat=51)  # xlOpenXMLWorkbook
        else:
            app.DisplayAlerts = 0
            update_links = app.Options.UpdateLinksAtOpen
            app.Options.UpdateLinksAtOpen = False
            document = app.Documents.Open(str(source), ReadOnly=True,
                                          AddToRecentFiles=False, ConfirmConversions=False,
                                          PasswordDocument="", WritePasswordDocument="")
            document.SaveAs2(str(output), FileFormat=12)  # wdFormatXMLDocument
    finally:
        failed = sys.exc_info()[0] is not None
        errors = []
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception as error:
                errors.append(str(error))
            document = None
        if app is not None:
            if update_links is not None:
                try:
                    app.Options.UpdateLinksAtOpen = update_links
                except Exception as error:
                    errors.append(str(error))
            try:
                _quit_application(application, app)
            except Exception as error:
                errors.append(str(error))
            app = None
        pythoncom.CoUninitialize()
        if errors and not failed:
            raise ValueError("Office cleanup failed: " + "; ".join(errors))


def run_office(operation, source, output, timeout=180):
    if operation not in APPLICATIONS:
        raise ValueError(f"Unknown Windows Office operation: {operation}")
    if platform.system() != "Windows":
        raise ValueError("Windows Office automation requires Windows")
    source, output = Path(source).resolve(), Path(output).resolve()
    if not source.is_file() or source == output or output.suffix.lower() != "." + operation:
        raise ValueError("Office needs an existing source and a separate output of the requested type")
    application = APPLICATIONS[operation]
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="coe-tos-win-office-") as directory:
        root = Path(directory)
        request, result, owner = (root / name for name in ("request.json", "result.json", "owner.json"))
        _write_json(request, {"operation": operation, "source": str(source), "output": str(output)})
        command = [sys.executable, str(WORKER), "--worker", "--request", str(request),
                   "--result", str(result), "--owner", str(owner)]
        try:
            process = subprocess.run(command, capture_output=True, text=True,
                                     encoding="utf-8", errors="replace", timeout=timeout,
                                     env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        except subprocess.TimeoutExpired as error:
            _, cleanup = _terminate_owned_process(owner)
            raise ValueError(f"Microsoft {application} timed out after {timeout}s. {cleanup}") from error
        except BaseException:
            _terminate_owned_process(owner)
            raise
        cleaned, cleanup = _terminate_owned_process(owner)
        try:
            info = json.loads(result.read_text(encoding="utf-8"))
        except (OSError, ValueError) as error:
            detail = process.stderr.strip() or "worker returned no valid result"
            raise ValueError(f"Microsoft {application} failed: {detail}. {cleanup}") from error
        if not isinstance(info, dict):
            raise ValueError(f"Microsoft {application} returned an invalid result. {cleanup}")
        if process.returncode or not info.get("ok"):
            detail = info.get("error") or process.stderr.strip() or "worker failed"
            raise ValueError(f"Microsoft {application} failed: {detail}. {cleanup}")
        if not cleaned:
            raise ValueError(f"Microsoft {application} cleanup incomplete: {cleanup}")
        if not output.is_file():
            raise ValueError(f"Microsoft {application} did not create {output.name}")
    return output


def main(argv=None):
    parser = argparse.ArgumentParser(description="Internal Windows Office worker")
    parser.add_argument("--worker", action="store_true", required=True)
    parser.add_argument("--request", required=True, type=Path)
    parser.add_argument("--result", required=True, type=Path)
    parser.add_argument("--owner", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        request = json.loads(args.request.read_text(encoding="utf-8"))
        operation = request["operation"]
        if operation not in APPLICATIONS:
            raise ValueError("Unknown Office operation")
        _perform(operation, Path(request["source"]), Path(request["output"]), args.owner)
        _write_json(args.result, {"ok": True})
        return 0
    except Exception as error:
        _write_json(args.result, {"ok": False, "error": f"{type(error).__name__}: {error}"})
        return 1


if __name__ == "__main__":
    sys.exit(main())

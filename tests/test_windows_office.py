"""Windows COM lifecycle tests run everywhere; real Office checks are opt-in."""
import contextlib
import io
import json
import os
from pathlib import Path
import platform
import sys
import tempfile
import time
import types
import unittest
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skills/coe-tos/scripts"))

import windows_office


class OperationTests(unittest.TestCase):
    def setUp(self):
        self.pythoncom = types.SimpleNamespace(CoInitialize=Mock(), CoUninitialize=Mock(), Missing=object())
        self.app, self.doc = Mock(), Mock()
        self.app.Workbooks.Open.return_value = self.doc
        self.app.Documents.Open.return_value = self.doc
        self.patches = [patch.dict(sys.modules, {"pythoncom": self.pythoncom}),
                        patch.object(windows_office, "_create_application", return_value=self.app)]
        for item in self.patches:
            item.start()
            self.addCleanup(item.stop)

    def perform(self, operation):
        windows_office._perform(operation, Path("source with spaces.xls"),
                                Path("out." + operation), Path("owner.json"))

    def test_pdf_export_respects_print_areas_and_closes_only_its_workbook(self):
        def export(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish):
            self.assertEqual(Type, 0)
            self.assertEqual(Quality, 0)
            self.assertFalse(IgnorePrintAreas)
            self.assertFalse(OpenAfterPublish)
            self.assertIs(From, self.pythoncom.Missing)
            self.assertIs(To, self.pythoncom.Missing)

        self.doc.ExportAsFixedFormat.side_effect = export
        self.perform("pdf")
        self.assertFalse(self.app.Visible)
        self.assertFalse(self.app.EnableEvents)
        self.assertEqual(self.app.AutomationSecurity, 3)
        self.assertEqual(self.app.Workbooks.Open.call_args.kwargs["UpdateLinks"], 0)
        self.assertTrue(self.app.Workbooks.Open.call_args.kwargs["ReadOnly"])
        self.app.CalculateFull.assert_called_once()
        self.doc.SaveAs.assert_not_called()
        self.doc.Close.assert_called_once_with(SaveChanges=0)
        self.app.Quit.assert_called_once()
        self.pythoncom.CoUninitialize.assert_called_once()

    def test_word_and_excel_use_correct_macro_free_conversion_formats(self):
        self.app.Options.UpdateLinksAtOpen = True
        self.perform("docx")
        self.doc.SaveAs2.assert_called_once_with(str(Path("out.docx")), FileFormat=12)
        self.app.Workbooks.Open.assert_not_called()
        self.assertTrue(self.app.Options.UpdateLinksAtOpen)
        self.app.Quit.assert_called_once_with(SaveChanges=0)
        self.app.reset_mock()
        self.doc.reset_mock()
        self.perform("xlsx")
        self.doc.SaveAs.assert_called_once_with(str(Path("out.xlsx")), FileFormat=51)
        self.assertFalse(self.doc.CheckCompatibility)
        self.app.CalculateFull.assert_not_called()

    def test_failure_during_open_still_quits_and_uninitializes_com(self):
        self.app.Workbooks.Open.side_effect = RuntimeError("Cannot open workbook")
        with self.assertRaisesRegex(RuntimeError, "Cannot open workbook"):
            self.perform("pdf")
        self.doc.Close.assert_not_called()
        self.app.Quit.assert_called_once()
        self.pythoncom.CoUninitialize.assert_called_once()

    def test_export_failure_survives_a_cleanup_failure(self):
        self.doc.ExportAsFixedFormat.side_effect = RuntimeError("PDF export failed")
        self.doc.Close.side_effect = RuntimeError("Close failed")
        with self.assertRaisesRegex(RuntimeError, "PDF export failed"):
            self.perform("pdf")
        self.app.Quit.assert_called_once()
        self.pythoncom.CoUninitialize.assert_called_once()

    def test_success_with_failed_cleanup_is_reported_as_incomplete(self):
        self.app.Quit.side_effect = RuntimeError("Quit failed")
        with self.assertRaisesRegex(ValueError, "Office cleanup failed.*Quit failed"):
            self.perform("pdf")


class OwnershipTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.owner = Path(temp.name) / "owner.json"
        self.app = Mock(Hwnd=77)
        self.client = types.ModuleType("win32com.client")
        self.client.DispatchEx = Mock(return_value=self.app)
        package = types.ModuleType("win32com")
        package.client = self.client
        self.process = Mock()
        self.process.EnumProcesses.return_value = [10, 20]
        self.process.GetWindowThreadProcessId.return_value = (50, 30)
        self.modules = {"win32com": package, "win32com.client": self.client,
                        "win32process": self.process}

    def test_creation_records_a_new_instance_and_refuses_a_preexisting_process(self):
        with patch.dict(sys.modules, self.modules), \
                patch.object(windows_office, "_creation_time", return_value="created-at-A"):
            self.assertIs(windows_office._create_application("Excel", self.owner), self.app)
            self.client.DispatchEx.assert_called_once_with("Excel.Application")
            self.assertEqual(json.loads(self.owner.read_text())["pid"], 30)
            self.owner.unlink()
            self.process.GetWindowThreadProcessId.return_value = (50, 20)
            with self.assertRaisesRegex(ValueError, "reused an existing process"):
                windows_office._create_application("Excel", self.owner)
            self.app.Quit.assert_not_called()
            self.assertFalse(self.owner.exists())

    def test_word_uses_window_hwnd_and_closes_its_bootstrap_document(self):
        bootstrap = Mock()
        bootstrap.Windows.Item.return_value.Hwnd = 88
        self.app.Documents.Add.return_value = bootstrap
        with patch.dict(sys.modules, self.modules), \
                patch.object(windows_office, "_creation_time", return_value="created-at-A"):
            windows_office._create_application("Word", self.owner)
        self.process.GetWindowThreadProcessId.assert_called_once_with(88)
        bootstrap.Close.assert_called_once_with(SaveChanges=0)

    def cleanup_modules(self, created):
        api, events, process = Mock(), Mock(), Mock()
        process.GetProcessTimes.return_value = {"CreationTime": created}
        events.WaitForSingleObject.side_effect = [258, 0]
        self.owner.write_text(json.dumps({"pid": 30, "created": "created-at-A"}), encoding="utf-8")
        return api, events, {"win32api": api, "win32event": events, "win32process": process}

    def test_timeout_cleanup_terminates_only_matching_pid_and_creation_time(self):
        api, events, modules = self.cleanup_modules("created-at-A")
        with patch.dict(sys.modules, modules):
            success, message = windows_office._terminate_owned_process(self.owner)
        self.assertTrue(success)
        self.assertIn("terminated", message)
        api.TerminateProcess.assert_called_once_with(api.OpenProcess.return_value, 1)
        api.OpenProcess.return_value.Close.assert_called_once()

    def test_reused_pid_is_never_terminated(self):
        api, events, modules = self.cleanup_modules("created-at-B")
        with patch.dict(sys.modules, modules):
            success, message = windows_office._terminate_owned_process(self.owner)
        self.assertTrue(success)
        self.assertIn("reused", message)
        api.TerminateProcess.assert_not_called()
        events.WaitForSingleObject.assert_not_called()

    def test_missing_ownership_record_does_not_attempt_process_cleanup(self):
        success, message = windows_office._terminate_owned_process(self.owner)
        self.assertFalse(success)
        self.assertIn("No verified Office PID", message)


class WatchdogTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root / "épreuve with spaces.xlsx"
        self.source.write_text("success", encoding="utf-8")
        self.output = self.root / "output with spaces.pdf"
        self.worker = self.root / "stub worker.py"
        self.worker.write_text('''import argparse, json, pathlib, sys, time
p = argparse.ArgumentParser()
p.add_argument('--worker', action='store_true')
for key in ('request', 'result', 'owner'): p.add_argument('--' + key)
a = p.parse_args()
r = json.loads(pathlib.Path(a.request).read_text(encoding='utf-8'))
mode = pathlib.Path(r['source']).read_text(encoding='utf-8')
pathlib.Path(a.owner).write_text('{}', encoding='utf-8')
if mode == 'hang': time.sleep(60)
if mode == 'fail':
    pathlib.Path(a.result).write_text(json.dumps({'ok': False, 'error': 'Échec export'}), encoding='utf-8')
    sys.exit(1)
pathlib.Path(r['output']).write_bytes(b'exported copy')
pathlib.Path(a.result).write_text(json.dumps({'ok': True}), encoding='utf-8')
''', encoding="utf-8")
        for item in (patch.object(windows_office.platform, "system", return_value="Windows"),
                     patch.object(windows_office, "WORKER", self.worker),
                     patch.object(windows_office, "_terminate_owned_process", return_value=(True, "cleaned"))):
            item.start()
            self.addCleanup(item.stop)

    def test_worker_receives_unicode_paths_and_uses_the_current_interpreter(self):
        with patch.object(windows_office.subprocess, "run", wraps=windows_office.subprocess.run) as run:
            self.assertEqual(windows_office.run_office("pdf", self.source, self.output, timeout=10), self.output.resolve())
            self.assertEqual(run.call_args.args[0][0], sys.executable)
        self.assertEqual(self.output.read_bytes(), b"exported copy")
        self.assertEqual(self.source.read_text(encoding="utf-8"), "success")

    def test_real_subprocess_timeout_is_bounded_and_preserves_existing_output(self):
        self.source.write_text("hang", encoding="utf-8")
        self.output.write_bytes(b"previous PDF")
        started = time.monotonic()
        with self.assertRaisesRegex(ValueError, "timed out.*cleaned"):
            windows_office.run_office("pdf", self.source, self.output, timeout=0.2)
        self.assertLess(time.monotonic() - started, 5)
        self.assertEqual(self.output.read_bytes(), b"previous PDF")

    def test_worker_failure_is_actionable_and_unicode_is_preserved(self):
        self.source.write_text("fail", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "Échec export"):
            windows_office.run_office("pdf", self.source, self.output, timeout=10)
        self.assertFalse(self.output.exists())


@unittest.skipUnless(platform.system() == "Windows" and os.environ.get("COE_TOS_RUN_OFFICE_TESTS") == "1",
                     "real Windows Office tests require explicit COE_TOS_RUN_OFFICE_TESTS=1")
class NativeOfficeTests(unittest.TestCase):
    def test_native_excel_export_leaves_an_unrelated_workbook_open(self):
        import pythoncom
        import win32com.client
        from coe_tos import main
        from export_pdf import export
        from pypdf import PdfReader

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with contextlib.redirect_stdout(io.StringIO()):
                main(["build", "--draft", str(ROOT / "examples/draft.json"), "--out", str(root), "--form-only"])
            original = (root / "TOS.xlsx").read_bytes()
            pythoncom.CoInitialize()
            app = document = None
            try:
                app = win32com.client.DispatchEx("Excel.Application")
                app.Visible, app.DisplayAlerts = False, False
                document = app.Workbooks.Add()
                name = document.Name
                info = export(root / "TOS.xlsx", root / "TOS.pdf", "excel-windows")
                self.assertEqual(info["pages"], 1)
                self.assertIn("TABLE OF SPECIFICATIONS", PdfReader(root / "TOS.pdf").pages[0].extract_text())
                self.assertEqual(document.Name, name)
                self.assertEqual((root / "TOS.xlsx").read_bytes(), original)
                # A visible mapping workbook exports its supplementary sheets too.
                with contextlib.redirect_stdout(io.StringIO()):
                    main(["build", "--draft", str(ROOT / "examples/draft.json"), "--out", str(root)])
                original = (root / "TOS.xlsx").read_bytes()
                info = export(root / "TOS.xlsx", root / "mapping.pdf", "excel-windows")
                self.assertGreater(info["pages"], 1)
                text = "\n".join(page.extract_text() for page in PdfReader(root / "mapping.pdf").pages)
                self.assertIn("Assessment mapping", text)
                self.assertIn("Allocation ledger", text)
                self.assertEqual(document.Name, name)
                self.assertEqual((root / "TOS.xlsx").read_bytes(), original)
            finally:
                if document is not None:
                    document.Close(SaveChanges=0)
                if app is not None:
                    app.Quit()
                pythoncom.CoUninitialize()

    def test_native_word_and_excel_legacy_inputs_are_extracted(self):
        import pythoncom
        import win32com.client
        from documents import extract

        with tempfile.TemporaryDirectory() as directory:
            pythoncom.CoInitialize()
            try:
                for application, suffix, format_code in (("Word", ".doc", 0), ("Excel", ".xls", 56)):
                    with self.subTest(application=application):
                        source = Path(directory) / ("épreuve legacy" + suffix)
                        app = win32com.client.DispatchEx(f"{application}.Application")
                        document = None
                        try:
                            app.Visible, app.DisplayAlerts = False, 0
                            if application == "Word":
                                document = app.Documents.Add()
                                document.Content.Text = "Q1: 50 points"
                                document.SaveAs2(str(source), FileFormat=format_code)
                            else:
                                document = app.Workbooks.Add()
                                document.Worksheets.Item(1).Range("A1").Value = "Q1: 50 points"
                                document.SaveAs(str(source), FileFormat=format_code)
                        finally:
                            if document is not None:
                                document.Close(SaveChanges=0)
                            app.Quit()
                        original = source.read_bytes()
                        result = extract(source, office_backend="ms-office")
                        self.assertTrue(any("Q1: 50" in row["text"] for row in result["segments"]))
                        self.assertEqual(source.read_bytes(), original)
            finally:
                pythoncom.CoUninitialize()


if __name__ == "__main__":
    unittest.main()

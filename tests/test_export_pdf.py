import importlib.util
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "coe-tos" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import export_pdf

HAS_PYMUPDF = importlib.util.find_spec("pymupdf") is not None


@unittest.skipUnless(HAS_PYMUPDF, "optional render tests require PyMuPDF")
class RenderTests(unittest.TestCase):
    def setUp(self):
        import pymupdf
        self.pymupdf = pymupdf
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.pdf = self.root / "form.pdf"
        self.write_pdf("first")

    def write_pdf(self, text):
        with self.pymupdf.open() as document:
            page = document.new_page()
            page.insert_text((40, 40), text)
            document.save(self.pdf)

    def test_cache_and_content_change(self):
        first = export_pdf.render_pdf(self.pdf, 1)
        second = export_pdf.render_pdf(self.pdf, 1)
        self.assertFalse(first["cache_hit"])
        self.assertTrue(second["cache_hit"])
        self.assertEqual(second["visual_review"], "pending")
        self.pdf.unlink()
        self.write_pdf("changed")
        changed = export_pdf.render_pdf(self.pdf, 1)
        self.assertFalse(changed["cache_hit"])
        self.assertNotEqual(first["renders"], changed["renders"])

    def test_corrupt_image_and_changed_scale(self):
        first = export_pdf.render_pdf(self.pdf, 1)
        Path(first["renders"][0]).write_bytes(b"corrupt")
        self.assertFalse(export_pdf.render_pdf(self.pdf, 1)["cache_hit"])
        self.assertNotEqual(first["renders"], export_pdf.render_pdf(self.pdf, 2)["renders"])

    def test_bad_scale(self):
        for scale in (0, -1, 5, float("nan"), float("inf")):
            with self.subTest(scale=scale), self.assertRaises(ValueError):
                export_pdf.render_pdf(self.pdf, scale)


class PreflightAndExcelTests(unittest.TestCase):
    def test_missing_renderer_fails_before_office(self):
        with patch.object(export_pdf, "renderer", side_effect=ValueError("missing renderer")), \
                patch.object(export_pdf.subprocess, "run") as run:
            with self.assertRaisesRegex(ValueError, "missing renderer"):
                export_pdf.export(Path("form.xlsx"), Path("out.pdf"), render=True)
            run.assert_not_called()

    def test_excel_owns_only_named_temporary_workbook_and_reports_stderr(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "form.xlsx"
            source.write_bytes(b"original workbook")
            output = root / "out.pdf"
            output.write_bytes(b"previous PDF")
            error = subprocess.CalledProcessError(1, ["osascript"], stderr="Excel parameter error -50")
            with patch.object(export_pdf, "preflight", return_value="excel-mac"), \
                    patch.object(export_pdf.subprocess, "run", side_effect=error) as run:
                with self.assertRaisesRegex(ValueError, "parameter error -50"):
                    export_pdf.export(source, output)
                script = run.call_args.kwargs["input"]
                self.assertIn("save workbook as targetWorkbook", script)
                self.assertIn("set targetWorkbook to workbook", script)
                self.assertNotIn("active workbook", script)
                self.assertNotIn("close every", script)
            self.assertEqual(source.read_bytes(), b"original workbook")
            self.assertEqual(output.read_bytes(), b"previous PDF")


if __name__ == "__main__":
    unittest.main()

import contextlib
import copy
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "coe-tos" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from coe_tos import main, report
from scoring import GROUPS, validate

ASSETS = ROOT / "skills" / "coe-tos" / "assets"
EXAMPLE = ROOT / "examples" / "draft.json"


class IngestTextTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)

    def ingest(self, text, name="exam-rubric.txt", append=None, out=None):
        out = out or self.directory / "sources.json"
        command = ["ingest-text", "--name", name, "--out", str(out), "--stdin"]
        if append:
            command += ["--append", str(append)]
        with contextlib.redirect_stdout(io.StringIO()), patch.object(sys, "stdin", io.StringIO(text)):
            status = main(command)
        return status, json.loads(out.read_text())

    def test_stdin_uses_line_locators_and_supplied_name(self):
        status, report_json = self.ingest("TAKE-HOME EXAM\n100 points\n\nCriterion\tPoints\n")
        self.assertEqual(status, 0)
        document = report_json["documents"][-1]
        self.assertEqual(document["source"], "exam-rubric.txt")
        self.assertEqual([segment["locator"] for segment in document["segments"]],
                         ["line 1", "line 2", "line 4"])  # blank line 3 is not a segment
        self.assertEqual(document["segments"][2]["text"], "Criterion\tPoints")
        self.assertTrue(any("chat" in warning for warning in document["warnings"]))

    def test_append_keeps_earlier_documents(self):
        first = self.directory / "first.json"
        self.ingest("Part one\n", name="exam.txt", out=first)
        _, merged = self.ingest("Part two\n", name="rubric.txt", append=first)
        self.assertEqual([document["source"] for document in merged["documents"]],
                         ["exam.txt", "rubric.txt"])

    def test_file_and_stdin_are_mutually_exclusive(self):
        target = self.directory / "out.json"
        both = ["ingest-text", "--name", "x.txt", str(EXAMPLE), "--out", str(target), "--stdin"]
        neither = ["ingest-text", "--name", "x.txt", "--out", str(target)]
        with contextlib.redirect_stderr(io.StringIO()):
            for command in (both, neither):
                with self.subTest(command=command):
                    with self.assertRaises(SystemExit):
                        main(command)
        self.assertFalse(target.exists())

    def test_empty_text_warns_instead_of_looking_readable(self):
        _, report_json = self.ingest("\n \n")
        warnings = report_json["documents"][-1]["warnings"]
        self.assertTrue(any("No readable content" in warning for warning in warnings))


class AssessmentMappingReportTests(unittest.TestCase):
    def setUp(self):
        self.draft = json.loads(EXAMPLE.read_text())
        self.profile = json.loads((ASSETS / "cjc-profile.json").read_text())
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.result = validate(self.draft, self.profile)
        self.text = report(self.draft, self.result, self.profile)

    def section(self, heading):
        body = self.text.split(f"## {heading}", 1)[1]
        return body.split("\n## ", 1)[0]

    def test_assessment_mapping_matches_published_scoring(self):
        self.assertEqual(self.result["errors"], [])
        section = self.section("Assessment mapping")
        self.assertIn("| Q1-C | C | 10 | 10 | 2 | 4 | 4 | Published subcriteria retained |", section)
        self.assertIn("| Q1-C / recall | recall | 2 | 2 | 2 | 0 | 0 | Published subcriterion |", section)
        self.assertIn("| Q1-A | A | 4 | 4 | 4 | 0 | 0 | Single allocation |", section)
        self.assertNotIn("Q1-C / recall | recall | 2 | 4", section)

    def test_criterion_titles_are_used_when_supplied(self):
        self.draft["questions"][0]["criteria"][0]["title"] = "Governing model and definitions"
        section = report(self.draft, self.result, self.profile).split("## Assessment mapping", 1)[1]
        self.assertIn("| Q1-A | Governing model and definitions | 4 | 4 |", section)

    def test_proposed_partition_is_labelled_for_existing_exams(self):
        existing = copy.deepcopy(self.draft)
        existing["mode"] = "existing-exam"
        criterion = existing["questions"][0]["criteria"][0]
        criterion["points"] = 8
        criterion.pop("subcriteria", None)
        first = existing["allocations"][0]
        second = copy.deepcopy(first)
        second.update(id="q1-a2", points=4, scores={"remembering": 0, "understanding": 4, "thinking": 0},
                      topic="visual", rationale="Applying the definition to this dataset.")
        existing["allocations"].append(second)
        section = report(existing, validate(existing, self.profile), self.profile)
        section = section.split("## Assessment mapping", 1)[1].split("\n## ", 1)[0]
        self.assertIn("Proposed topic partition", section)

    def test_topic_allocation_totals_reconcile(self):
        section = self.section("Topic allocation")
        rows = [line for line in section.splitlines() if line.startswith("| ") and "TOTAL" not in line]
        self.assertEqual(len(rows), len(self.draft["topics"]) + 1)  # header row plus one row per topic
        totals = [line for line in section.splitlines() if line.startswith("| TOTAL")]
        self.assertEqual(len(totals), 1)
        cells = [cell.strip() for cell in totals[0].strip("|").split("|")]
        expected = [str(self.result["cognitive"][group]) for group in GROUPS]
        self.assertEqual(cells[1:4], expected)
        self.assertEqual(cells[4], str(self.result["total"]))

    def test_band_ranges_are_shown(self):
        self.assertIn("| Remembering | 16 | 16.00% | Within band (10-20%) |", self.section("Totals"))
        self.assertIn("| Thinking | 60 | 60.00% | Within band (60-100%) |", self.section("Totals"))

    def test_source_inputs_and_chat_provenance(self):
        self.draft["metadata"]["input_mode"] = "chat-pasted"
        self.draft["metadata"]["totals_confirmed_by"] = "requester"
        self.draft["metadata"]["totals_confirmed_on"] = "2026-10-06"
        section = report(self.draft, validate(self.draft, self.profile), self.profile)
        section = section.split("## Source inputs", 1)[1].split("\n## ", 1)[0]
        self.assertIn("- input_mode: chat-pasted", section)
        self.assertIn("- Published totals confirmed by: requester on 2026-10-06", section)
        self.assertIn("| assessment.md | 24 | 12 |", section)

    def test_chat_pasted_input_warns_without_blocking_the_build(self):
        self.draft["metadata"]["input_mode"] = "chat-pasted"
        result = validate(self.draft, self.profile)
        self.assertEqual(result["errors"], [])
        self.assertTrue(any("chat-pasted" in warning for warning in result["warnings"]))
        self.draft["metadata"]["input_mode"] = "file"
        self.assertFalse(any("chat-pasted" in warning for warning in validate(self.draft, self.profile)["warnings"]))

    def test_generated_review_file_contains_new_sections(self):
        output = self.directory / "out"
        self.assertEqual(main(["build", "--draft", str(EXAMPLE), "--out", str(output), "--xlsx-only"]), 0)
        text = (output / "review.md").read_text(encoding="utf-8")
        for heading in ("Assessment mapping", "Topic allocation", "Source inputs"):
            self.assertIn(f"## {heading}", text)

    def test_cjc_semester_list_leading_space_is_exact(self):
        # The bundled list stores " 1st" with a leading space but "2nd"/"Summer" without one.
        for value, accepted in ((" 2nd", False), ("2nd", True)):
            draft = copy.deepcopy(self.draft)
            draft["metadata"]["semester"] = value
            draft_path = self.directory / f"draft-{value.strip()}.json"
            draft_path.write_text(json.dumps(draft, indent=2), encoding="utf-8")
            target = self.directory / f"out-{value.strip()}"
            with contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                try:
                    status = main(["build", "--draft", str(draft_path), "--out", str(target), "--xlsx-only"])
                except SystemExit as exit_error:
                    status = exit_error.code
            self.assertEqual(status == 0, accepted, value)
            self.assertEqual(target.joinpath("TOS.xlsx").exists(), accepted, value)


if __name__ == "__main__":
    unittest.main()

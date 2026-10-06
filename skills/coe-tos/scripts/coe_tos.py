#!/usr/bin/env python3
"""CoE-TOS portable CLI: extract -> inspect -> validate -> build -> export."""
import argparse
import csv
import json
from pathlib import Path
import sys
import tempfile
import time
from decimal import Decimal

from export_pdf import export, preflight, render_pdf
from scoring import GROUPS, validate

SKILL = Path(__file__).resolve().parents[1]
ASSETS = SKILL / "assets"


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False,
                               default=lambda x: str(x) if isinstance(x, Decimal) else str(x)) + "\n", encoding="utf-8")


def report(draft, result, profile=None, fidelity=None, pdf=None):
    def safe(value):
        return str(value).replace("|", "\\|").replace("\n", " ")

    def points(value):
        """Render a score without float artifacts such as 12.000000000000002."""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    bands = (profile or {}).get("bands", {})

    def band_text(group):
        bounds = bands.get(group)
        if not bounds:
            return "Not specified"
        shown = "-".join(points(value) for value in bounds)
        return ("Within" if result["bands"].get(group) else "Outside") + f" band ({shown}%)"

    def ledger_totals(entries):
        ledger = sum((Decimal(str(entry.get("points", 0))) for entry in entries), Decimal(0))
        scores = [sum((Decimal(str(entry.get("scores", {}).get(group, 0))) for entry in entries), Decimal(0))
                  for group in GROUPS]
        return ledger, scores

    metadata = draft.get("metadata", {})
    lines = ["# CoE-TOS review", "", f"Course: {metadata.get('course_title', '[To be supplied]')}",
             f"Mode: {draft.get('mode')}", "", "## Status", "",
             "Arithmetic checks: " + ("FAILED" if result["errors"] else "PASSED"),
             "Cognitive classifications: evidence-led draft; instructor review required.",
             "Source citations are supplied by the agent; arithmetic validation does not independently verify their meaning.",
             "", "## Totals", "", "| Group | Points | Percent | Band |", "|:--|--:|--:|:--|"]
    for group in GROUPS:
        lines.append(f"| {group.title()} | {result['cognitive'][group]} | {result['shares'][group]:.2f}% | {band_text(group)} |")
    lines += [f"| Total | {result['total']} | {'100%' if result['total'] else 'Undefined'} | |"]

    lines += ["", "## Assessment mapping", "",
              "Published scoring mapped to the allocation ledger. `Ref` matches the workbook Test Item No. cells.",
              "Source points come from the rubric; ledger points are what the allocations award.",
              "", "| Ref | Published criterion | Source points | Ledger points | R | U | T | Partition |",
              "|:--|:--|--:|--:|--:|--:|--:|:--|"]
    grouped = {}
    for allocation in draft.get("allocations", []):
        grouped.setdefault((allocation.get("question"), allocation.get("criterion")), []).append(allocation)
    for question in draft.get("questions", []):
        for criterion in question.get("criteria", []):
            entries = grouped.get((question.get("id"), criterion.get("id")), [])
            ledger, scores = ledger_totals(entries)
            declared = criterion.get("subcriteria") or {}
            partition = ("Published subcriteria retained" if declared else
                         "Proposed topic partition" if len(entries) > 1 else "Single allocation")
            lines.append("| " + " | ".join(map(safe, [
                f"{question.get('id')}-{criterion.get('id')}",
                criterion.get("title") or criterion.get("id"),
                points(criterion.get("points")), points(ledger),
                *(points(score) for score in scores), partition])) + " |")
            for sub, sub_points in declared.items():
                sub_ledger, sub_scores = ledger_totals([entry for entry in entries if entry.get("subcriterion") == sub])
                lines.append("| " + " | ".join(map(safe, [
                    f"{question.get('id')}-{criterion.get('id')} / {sub}", str(sub),
                    points(sub_points), points(sub_ledger),
                    *(points(score) for score in sub_scores), "Published subcriterion"])) + " |")

    lines += ["", "## Topic allocation", "", "| Topic | R | U | T | Total | Share |", "|:--|--:|--:|--:|--:|--:|"]
    rows = []
    for topic in draft.get("topics", []):
        scores = result["topic_scores"].get(topic.get("id"), {})
        values = [Decimal(str(scores.get(group, 0))) for group in GROUPS]
        total = sum(values, Decimal(0))
        rows.append(values)
        share = f"{total * 100 / result['total']:.2f}%" if result["total"] else "n/a"
        lines.append("| " + " | ".join(map(safe, [topic.get("title") or topic.get("id"),
                                                   *(points(value) for value in values),
                                                   points(total), share])) + " |")
    totals = [sum((row[index] for row in rows), Decimal(0)) for index in range(len(GROUPS))]
    grand = sum(totals, Decimal(0))
    share = f"{grand * 100 / result['total']:.2f}%" if result["total"] else "n/a"
    lines.append("| " + " | ".join(["TOTAL", *(points(value) for value in totals), points(grand), share]) + " |")

    lines += ["", "## Question reconciliation", "",
              "| Question | Remembering | Understanding | Thinking | Total |", "|:--|--:|--:|--:|--:|"]
    for question, scores in result["question_scores"].items():
        lines.append(f"| {safe(question)} | " + " | ".join(str(scores[g]) for g in GROUPS) + f" | {sum(scores.values())} |")
    lines += ["", "## Allocation ledger", "",
              "| Allocation | Question / criterion / subcriterion | Topic | Points | R / U / T | Evidence | Rationale |",
              "|:--|:--|:--|--:|:--|:--|:--|"]
    for allocation in draft.get("allocations", []):
        evidence = "; ".join(f"{e.get('source')}: {e.get('locator')}" for e in allocation.get("evidence", []))
        ref = f"{allocation.get('question', '?')}/{allocation.get('criterion', '?')}/{allocation.get('subcriterion', '—')}"
        scores = " / ".join(str(allocation.get("scores", {}).get(g, "?")) for g in GROUPS)
        lines.append("| " + " | ".join(map(safe, [allocation.get("id", "?"), ref, allocation.get("topic", "?"), allocation.get("points", "?"), scores, evidence, allocation.get("rationale", "")])) + " |")
    provenance = {}
    for allocation in draft.get("allocations", []):
        for item in allocation.get("evidence", []):
            entry = provenance.setdefault(item.get("source", "?"), {"allocations": set(), "locators": set()})
            entry["allocations"].add(allocation.get("id"))
            entry["locators"].add(item.get("locator"))
    lines += ["", "## Source inputs", ""]
    notes = []
    if metadata.get("input_mode"):
        notes.append(f"- input_mode: {metadata['input_mode']}")
    if metadata.get("totals_confirmed_by"):
        confirmed = f"- Published totals confirmed by: {metadata['totals_confirmed_by']}"
        if metadata.get("totals_confirmed_on"):
            confirmed += f" on {metadata['totals_confirmed_on']}"
        notes.append(confirmed)
    lines.extend(notes)
    if provenance:
        lines += [""] if notes else []
        lines += ["| Source | Allocations | Distinct locators |", "|:--|--:|--:|"]
        for source, entry in sorted(provenance.items()):
            lines.append(f"| {safe(source)} | {len(entry['allocations'])} | {len(entry['locators'])} |")
    lines.append(f"- Distinct evidence locators recorded: {result.get('evidence_locators', 0)}")

    lines += ["", "## Checks and unresolved items", ""]
    lines.extend("- ERROR: " + e for e in result["errors"])
    lines.extend("- " + w for w in result["warnings"])
    if not result["errors"]:
        lines += ["- Question, criterion and declared subcriterion totals reconcile; scores are allocated once in the ledger."]
    if fidelity:
        lines += [f"- Template SHA-256: `{fidelity['template_sha256']}`.",
                  f"- All {fidelity['package_parts_preserved']} package parts retained; original drawing/media parts are byte-identical.",
                  f"- {fidelity['formula_count']} original formulas retained; styles, merges, validations and protection preserved."]
        if fidelity.get("added_sheets"):
            lines += ["- The template's own worksheet changed only in declared values, formula caches and print settings.",
                      f"- Supplementary worksheets added: {', '.join(fidelity['added_sheets'])}. They are hidden, so the PDF export contains the form only; unhide them in Excel to read them."]
        else:
            lines.append("- Explicit profile print-area/fit settings and calculation flags are the only non-value changes.")
        if fidelity["unsupported_formula_caches"]:
            lines += ["- Some formula caches require Excel/LibreOffice recalculation:"]
            lines.extend("  - " + w for w in fidelity["unsupported_formula_caches"])
    if pdf:
        if pdf.get("error"):
            lines += ["- PDF export incomplete: " + pdf["error"]]
        else:
            lines += [f"- PDF exported with {pdf['backend']}: {pdf['pages']} pages, searchable text.",
                      "- Visual review pending: inspect every page for header/logo presence, clipping, band highlights and page breaks."]
    else:
        lines += ["- PDF export not performed."]
    lines += ["", "## Metadata", ""]
    lines.extend(f"- {key}: {value}" for key, value in metadata.items())
    lines += ["", "## Assessment changes", ""]
    lines.extend("- " + str(change) for change in draft.get("proposed_changes", []))
    if not draft.get("proposed_changes"):
        lines.append("No proposed scoring changes recorded.")
    return "\n".join(lines) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    command = commands.add_parser("extract", help="Extract multiple sources to JSON with locators")
    command.add_argument("inputs", nargs="+", type=Path)
    command.add_argument("--out", required=True, type=Path)
    command.add_argument("--ocr", action="store_true")
    command = commands.add_parser("ingest-text", help="Add text supplied in chat to a sources JSON report")
    command.add_argument("--name", required=True, help="Label recorded as the source, for example exam-rubric.txt")
    command.add_argument("--stdin", action="store_true", help="Read the supplied text from standard input")
    command.add_argument("input", nargs="?", type=Path, help="Text file to ingest; omit when using --stdin")
    command.add_argument("--append", type=Path, help="Existing sources JSON to merge into")
    command.add_argument("--out", required=True, type=Path)
    command = commands.add_parser("inspect-template", help="Inspect fields, formulas, validations, dimensions and drawings")
    command.add_argument("template", nargs="?", type=Path, default=ASSETS / "cjc-template.xlsx")
    command.add_argument("--out", required=True, type=Path)
    for name in ("validate", "build"):
        command = commands.add_parser(name)
        command.add_argument("--draft", required=True, type=Path)
        command.add_argument("--profile", type=Path, default=ASSETS / "cjc-profile.json")
        command.add_argument("--out", required=True, type=Path,
                             help="JSON validation report for validate; output directory for build")
        if name == "build":
            command.add_argument("--template", type=Path, default=ASSETS / "cjc-template.xlsx")
            command.add_argument("--xlsx-only", action="store_true")
            command.add_argument("--backend", choices=["auto", "excel-mac", "libreoffice"], default="auto")
            command.add_argument("--render", action="store_true")
            command.add_argument("--render-scale", type=float, default=1.5)
            command.add_argument("--with-mapping", action="store_true",
                                 help="Also add assessment mapping, topic allocation, ledger and notes sheets")
            command.add_argument("--mapping-visible", action="store_true",
                                 help="Show the supplementary sheets instead of hiding them from the PDF export")
    command = commands.add_parser("export-pdf")
    command.add_argument("workbook", type=Path)
    command.add_argument("--out", required=True, type=Path)
    command.add_argument("--backend", choices=["auto", "excel-mac", "libreoffice"], default="auto")
    command.add_argument("--render", action="store_true")
    command.add_argument("--render-scale", type=float, default=1.5)
    command = commands.add_parser("render-pdf", help="Render existing PDF page images without rebuilding or opening Office")
    command.add_argument("pdf", type=Path)
    command.add_argument("--scale", type=float, default=1.5, help="0.5 to 4; use 1 for quick previews")
    command = commands.add_parser("preview-template", help="Native PDF of the blank template using the profile's print settings")
    command.add_argument("--template", type=Path, default=ASSETS / "cjc-template.xlsx")
    command.add_argument("--profile", type=Path, default=ASSETS / "cjc-profile.json")
    command.add_argument("--out", required=True, type=Path)
    command.add_argument("--backend", choices=["auto", "excel-mac", "libreoffice"], default="auto")
    command.add_argument("--render", action="store_true")
    command.add_argument("--render-scale", type=float, default=1.5)
    args = parser.parse_args(argv)
    try:
        if args.command == "extract":
            from documents import extract
            documents = []
            failed = False
            for path in args.inputs:
                try:
                    documents.append(extract(path, ocr=args.ocr))
                except Exception as error:
                    failed = True
                    documents.append({"source": path.name, "error": str(error), "segments": []})
            write_json(args.out, {"version": 1, "documents": documents})
            return 1 if failed else 0
        if args.command == "ingest-text":
            from documents import text_document
            if args.stdin == bool(args.input):
                raise ValueError("Provide exactly one of --stdin or a text file path")
            text = sys.stdin.read() if args.stdin else args.input.read_text(encoding="utf-8")
            document = text_document(text, args.name)
            collected = {"version": 1, "documents": []}
            if args.append and args.append.is_file():
                collected["documents"].extend(read_json(args.append).get("documents", []))
            collected["documents"].append(document)
            write_json(args.out, collected)
            print(f"Ingested {len(document['segments'])} lines as {args.name} into {args.out}")
            return 0
        if args.command == "inspect-template":
            from workbooks import inspect
            write_json(args.out, inspect(args.template))
            return 0
        if args.command == "export-pdf":
            info = export(args.workbook, args.out, args.backend, args.render, args.render_scale)
            write_json(args.out.with_suffix(".export.json"), info)
            print(args.out)
            return 0
        if args.command == "render-pdf":
            info = render_pdf(args.pdf, args.scale)
            print(json.dumps(info, indent=2))
            return 0
        if args.command == "preview-template":
            from workbooks import preview_workbook
            preflight(args.backend, args.render, args.render_scale)
            with tempfile.TemporaryDirectory(prefix="coe-tos-preview-") as directory:
                copied = Path(directory) / "blank.xlsx"
                preview_workbook(args.template, copied, read_json(args.profile))
                info = export(copied, args.out, args.backend, args.render, args.render_scale)
            write_json(args.out.with_suffix(".export.json"), info)
            print(args.out)
            return 0
        started = time.perf_counter()
        draft, profile = read_json(args.draft), read_json(args.profile)
        result = validate(draft, profile)
        if args.command == "validate":
            write_json(args.out, result)
            print("FAILED" if result["errors"] else "PASSED", *result["errors"], sep="\n")
            return 1 if result["errors"] else 0
        args.out.mkdir(parents=True, exist_ok=True)
        write_json(args.out / "validation.json", result)
        if result["errors"]:
            (args.out / "review.md").write_text(report(draft, result, profile), encoding="utf-8")
            raise ValueError("Invalid draft: " + "; ".join(result["errors"]))
        if not args.xlsx_only:
            preflight(args.backend, args.render, args.render_scale)
        from workbooks import build
        extra = None
        if args.with_mapping:
            from mapping import sheets
            extra = sheets(draft, result)
        fidelity = build(args.template, args.out / "TOS.xlsx", draft, profile, result,
                         extra_sheets=extra, hidden_sheets=not args.mapping_visible)
        write_json(args.out / "draft.json", draft)
        write_json(args.out / "template-profile.json", profile)
        write_json(args.out / "fidelity.json", fidelity)
        pdf = None
        if not args.xlsx_only:
            try:
                pdf = export(args.out / "TOS.xlsx", args.out / "TOS.pdf", args.backend, args.render, args.render_scale)
            except Exception as error:
                pdf = {"error": str(error), "current_pdf_verified": False,
                       "warning": "Any pre-existing PDF in this output directory may be stale; do not deliver it."}
            write_json(args.out / "pdf-checks.json", pdf)
        (args.out / "review.md").write_text(report(draft, result, profile, fidelity, pdf), encoding="utf-8")
        with (args.out / "mapping.csv").open("w", newline="", encoding="utf-8") as stream:
            writer = csv.writer(stream)
            writer.writerow(["id", "question", "criterion", "subcriterion", "topic", "points", *GROUPS, "rationale"])
            for a in draft["allocations"]:
                writer.writerow([a["id"], a["question"], a["criterion"], a.get("subcriterion", ""), a["topic"], a["points"],
                                 *(a["scores"][g] for g in GROUPS), a["rationale"]])
        print("Created", args.out / "TOS.xlsx")
        write_json(args.out / "timings.json", {"total_seconds": round(time.perf_counter() - started, 4),
                   "export_seconds": (pdf or {}).get("export_seconds"),
                   "render_seconds": (pdf or {}).get("render_seconds")})
        print("Cognitive totals:", " / ".join(str(result["cognitive"][g]) for g in GROUPS))
        if pdf and pdf.get("error"):
            print("PDF incomplete:", pdf["error"], file=sys.stderr)
            return 1
        return 0
    except (ValueError, KeyError, OSError, TypeError) as error:
        parser.exit(1, f"CoE-TOS: {error}\n")


if __name__ == "__main__":
    sys.exit(main())

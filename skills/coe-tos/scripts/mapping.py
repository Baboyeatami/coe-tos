"""Optional supplementary worksheets: assessment mapping, topic allocation, ledger, notes.

These sheets are generated for the reader and attached directly to the OOXML
package. The institutional sheet keeps its structure/formulas/styles outside
declared edits; styles.xml, drawings and media retain their original bytes.
Supplementary tables are validated snapshots with no new formulas. Percentages
are preformatted text because the template's style table is preserved untouched.
"""
from decimal import Decimal
import json
import math

from lxml import etree as ET

from scoring import GROUPS, allocation_index
from openpyxl.utils import get_column_letter

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda name: f"{{{NS}}}{name}"
QCT = lambda name: f"{{{CT}}}{name}"
QPKG = lambda name: f"{{{PKG}}}{name}"
QREL = lambda name: f"{{{REL}}}{name}"
COLUMN = lambda index: get_column_letter(index + 1)


def text(value):
    return value if isinstance(value, str) else str(value)


def as_row(values, row_number):
    """Normalise one row to a list of cells.

    A bare string counts as a single cell. Anything else that is not a sequence is
    refused, because iterating it by accident emits malformed cell references.
    """
    if isinstance(values, str):
        return [values]
    if not isinstance(values, (list, tuple)):
        raise ValueError(f"Row {row_number} must be a list of cells, got {type(values).__name__}")
    return list(values)


def cells_of(values, row_number):
    """Yield (address, value) for one normalised row, skipping empty cells."""
    for index, value in enumerate(values):
        if value is not None:
            yield f"{COLUMN(index)}{row_number}", value


def template_styles(workbook):
    """Reuse existing general-format header/wrap styles without changing styles.xml."""
    result = {}
    for index, style in enumerate(workbook._cell_styles):
        if style.numFmtId != 0:
            continue
        alignment = workbook._alignments[style.alignmentId]
        font = workbook._fonts[style.fontId]
        if font.bold and "header" not in result:
            result["header"] = index
        if alignment.wrapText and not font.bold and "body" not in result:
            result["body"] = index
    return result


def worksheet_xml(rows, widths, styles=None):
    styles = styles or {}
    root = ET.Element(Q("worksheet"), nsmap={None: NS})
    # These sheets are readable evidence, so give them a landscape, fit-to-width page
    # setup; without it Excel splits the columns across many pages when printed.
    prop = ET.SubElement(root, Q("sheetPr"))
    ET.SubElement(prop, Q("pageSetUpPr"), fitToPage="1")
    views = ET.SubElement(root, Q("sheetViews"))
    view = ET.SubElement(views, Q("sheetView"), workbookViewId="0")
    ET.SubElement(view, Q("pane"), ySplit="5", topLeftCell="A6", activePane="bottomLeft", state="frozen")
    if widths:
        columns = ET.SubElement(root, Q("cols"))
        for index, width in enumerate(widths, 1):
            if width:
                ET.SubElement(columns, Q("col"), min=str(index), max=str(index),
                              width=str(width), customWidth="1")
    data = ET.SubElement(root, Q("sheetData"))
    for offset, values in enumerate(rows, 1):
        values = as_row(values, offset)
        if not any(value is not None for value in values):
            continue
        row = ET.SubElement(data, Q("row"), r=str(offset))
        if offset >= 5 and "body" in styles:
            lines = max((sum(max(1, math.ceil(len(line) / max(1, widths[i] - 2)))
                             for line in str(value).split("\n"))
                         for i, value in enumerate(values) if value is not None and i < len(widths)), default=1)
            row.set("ht", str(min(409, max(18, lines * 15))))
            row.set("customHeight", "1")
        for address, value in cells_of(values, offset):
            cell = ET.SubElement(row, Q("c"), r=address)
            style = styles.get("header") if offset == 5 else styles.get("body") if offset > 5 and isinstance(value, str) else None
            if style is not None:
                cell.set("s", str(style))
            if isinstance(value, bool):
                cell.set("t", "b")
                ET.SubElement(cell, Q("v")).text = str(int(value))
            elif isinstance(value, (int, float, Decimal)):
                ET.SubElement(cell, Q("v")).text = str(value)
            else:
                cell.set("t", "inlineStr")
                node = ET.SubElement(ET.SubElement(cell, Q("is")), Q("t"))
                node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                node.text = text(value)
    ET.SubElement(root, Q("pageMargins"), left="0.4", right="0.4", top="0.5", bottom="0.5",
                  header="0.3", footer="0.3")
    ET.SubElement(root, Q("pageSetup"), orientation="landscape", fitToWidth="1", fitToHeight="0")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _totals(entries):
    points = sum((Decimal(str(entry.get("points", 0))) for entry in entries), Decimal(0))
    scores = [sum((Decimal(str(entry["scores"].get(group, 0))) for entry in entries), Decimal(0))
              for group in GROUPS]
    return points, scores


def _plain(value):
    return text(int(value) if Decimal(str(value)) == Decimal(str(value)).to_integral_value() else value)


def assessment_mapping(draft, result, index=None):
    grouped = (index or allocation_index(draft))["criteria"]
    rows = [["Assessment mapping"], ["Published scoring beside the allocation ledger. Ref matches the Test Item No. cells on the form."],
            [f"Mode: {draft['mode']} | Total: {_plain(result['total'])} points | R/U/T: "
             + " / ".join(_plain(result["cognitive"][group]) for group in GROUPS)], [],
            ["Ref", "Published criterion", "Source points", "Ledger points", "R", "U", "T", "Partition"]]
    published = ledger = remembers = understands = thinks = Decimal(0)
    for question in draft["questions"]:
        for criterion in question["criteria"]:
            entries = grouped.get((question["id"], criterion["id"]), [])
            points, scores = _totals(entries)
            declared = criterion.get("subcriteria") or {}
            partition = ("Published subcriteria retained" if declared else
                         "Proposed topic partition" if len(entries) > 1 else "Single allocation")
            rows.append([f"{question['id']}-{criterion['id']}", criterion.get("title") or criterion["id"],
                         criterion["points"], points, scores[0], scores[1], scores[2], partition])
            published += Decimal(str(criterion["points"]))
            ledger += points
            remembers, understands, thinks = remembers + scores[0], understands + scores[1], thinks + scores[2]
            for sub, sub_points in declared.items():
                sub_ledger, sub_scores = _totals([e for e in entries if e.get("subcriterion") == sub])
                rows.append([f"{question['id']}-{criterion['id']} / {sub}", sub,
                             sub_points, sub_ledger, sub_scores[0], sub_scores[1], sub_scores[2],
                             "Published subcriterion"])
    rows.append(["TOTAL", "", published, ledger, remembers, understands, thinks, ""])
    return "Assessment mapping", rows, [26, 40, 13, 13, 7, 7, 7, 26]


def topic_allocation(draft, result, index=None):
    lookup = index or allocation_index(draft)
    rows = [["Topic allocation"], ["Cognitive points and share by topic, matching the contents grid on the form."],
            [f"Total: {_plain(result['total'])} points"],
            ["Refs column lists every criterion mapped to the topic; the form abbreviates when a cell cannot hold them all."],
            ["Topic", "R", "U", "T", "Total", "Share", "Refs"]]
    totals = [Decimal(0)] * 4
    for topic in draft["topics"]:
        scores = result["topic_scores"].get(topic["id"], {})
        values = [Decimal(str(scores.get(group, 0))) for group in GROUPS]
        total = sum(values, Decimal(0))
        share = f"{total * 100 / result['total']:.2f}%" if result["total"] else "n/a"
        refs = sorted({f"{a['question']}-{a['criterion']}" for a in lookup["topics"][topic["id"]]})
        rows.append([topic["title"], values[0], values[1], values[2], total, share, "; ".join(refs)])
        for index, value in enumerate(values):
            totals[index] += value
        totals[3] += total
    rows.append(["TOTAL", totals[0], totals[1], totals[2], totals[3],
                 f"{totals[3] * 100 / result['total']:.2f}%" if result["total"] else "n/a", ""])
    return "Topic allocation", rows, [46, 7, 7, 7, 9, 11, 34]


def allocation_ledger(draft):
    rows = [["Allocation ledger"], ["Every allocated score, allocated once. Source evidence is retained here; assumptions and checks are on Notes."],
            [f"Allocations: {len(draft['allocations'])}"], [],
            ["Allocation", "Question", "Criterion", "Subcriterion", "Topic", "Points", "R", "U", "T", "Rationale", "Sources", "Source locators"]]
    total = Decimal(0)
    for allocation in draft["allocations"]:
        total += Decimal(str(allocation["points"]))
        rows.append([allocation["id"], allocation["question"], allocation["criterion"],
                     allocation.get("subcriterion", ""), allocation["topic"], allocation["points"],
                     allocation["scores"]["remembering"], allocation["scores"]["understanding"],
                     allocation["scores"]["thinking"], allocation.get("rationale", ""),
                     "; ".join(dict.fromkeys(e["source"] for e in allocation["evidence"])),
                     "\n".join(f"{e['source']}: {e['locator']}" for e in allocation["evidence"])])
    rows.append(["TOTAL", "", "", "", "", total, "", "", "", "", "", ""])
    return "Allocation ledger", rows, [22, 9, 10, 16, 12, 8, 6, 6, 6, 70, 44, 85]


def notes(draft, result, profile=None):
    rows = [["Notes and unresolved items"],
            ["Arithmetic validation does not establish educational validity or source accuracy."], [], [],
            ["Kind", "Text"]]
    for key, value in draft.get("metadata", {}).items():
        if value is not None:
            shown = json.dumps(value, ensure_ascii=False) if isinstance(value, (list, dict)) else str(value)
            rows.append([f"Metadata: {key}", shown])
    rows.append(["Assessment mode", draft["mode"]])
    rows.append(["Exam points", result["total"]])
    rows.append(["Arithmetic validation", "Passed: published question, criterion and declared subcriterion totals reconcile."])
    rows.append(["Output visual review", "Not performed by this workbook-only build. Structural checks are not visual approval."])
    for group in GROUPS:
        bounds = (profile or {}).get("bands", {}).get(group)
        band = f"{bounds[0]}-{bounds[1]}%" if bounds else "not specified"
        status = ("within band" if result["bands"].get(group) else "outside band") if bounds else "no target band"
        rows.append([f"Cognitive: {group}", f"{result['cognitive'][group]} points; {result['shares'][group]:.2f}%; target {band}; {status}"])
    for text_value in draft.get("assumptions", []):
        rows.append(["Assumption", text_value])
    for text_value in draft.get("proposed_changes", []):
        rows.append(["Proposed change", text_value])
    for text_value in result.get("warnings", []):
        if text_value in draft.get("assumptions", []):
            continue
        rows.append(["Warning", text_value])
    rows.append(["Scoring accountability", "Cognitive classifications and any proposed within-criterion splits require instructor review; band conflicts are not repaired by relabelling tasks."])
    return "Notes", rows, [18, 110]


def sheets(draft, result, profile=None, index=None):
    """Ordered (name, rows, widths) specifications for the supplementary sheets."""
    index = index or allocation_index(draft)
    return [assessment_mapping(draft, result, index), topic_allocation(draft, result, index),
            allocation_ledger(draft), notes(draft, result, profile)]


def _next_id(rels):
    used = {relationship.get("Id") for relationship in rels}
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def attach(parts, specifications, hidden=True, styles=None):
    """Add worksheet parts to the package and register them in workbook.xml, rels and content types.

    Returns the list of new parts so the fidelity gate can account for them.
    """
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    sheet_nodes = workbook.find(Q("sheets"))
    existing = sheet_nodes.findall(Q("sheet"))
    used_names = {node.get("name") for node in existing}
    used_ids = {int(node.get("sheetId")) for node in existing}
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    content = ET.fromstring(parts["[Content_Types].xml"])
    added, relationship_ids, sheet_names = [], [], []
    for offset, (name, rows, widths) in enumerate(specifications, 1):
        name = name if name not in used_names else f"{name} {offset}"
        relationship = _next_id(rels)
        sheet_id = max(used_ids) + offset
        part = f"xl/worksheets/sheet{max(len(existing) + offset, 1)}.xml"
        index = offset
        while part in parts or part in added:
            index += 1
            part = f"xl/worksheets/sheet{max(len(existing) + offset + index, 1)}.xml"
        parts[part] = worksheet_xml(rows, widths, styles)
        added.append(part)
        relationship_ids.append(relationship)
        sheet_names.append(name)
        node = ET.SubElement(sheet_nodes, Q("sheet"), name=name, sheetId=str(sheet_id))
        node.set(QREL("id"), relationship)
        if hidden:
            node.set("state", "hidden")
        ET.SubElement(rels, QPKG("Relationship"), Id=relationship,
                      Type=f"{REL}/worksheet", Target=part.replace("xl/", ""))
        ET.SubElement(content, QCT("Override"), PartName=f"/{part}",
                      ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
    parts["xl/_rels/workbook.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    parts["[Content_Types].xml"] = ET.tostring(content, encoding="utf-8", xml_declaration=True)
    return {"parts": added, "relationship_ids": relationship_ids, "sheet_names": sheet_names}

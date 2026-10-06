"""Validate an evidence-led allocation ledger with exact decimal arithmetic."""
from collections import defaultdict
from decimal import Decimal, InvalidOperation

GROUPS = ("remembering", "understanding", "thinking")


def number(value):
    if isinstance(value, bool):
        raise ValueError("Boolean is not a score")
    try:
        result = Decimal(str(value))
    except InvalidOperation as error:
        raise ValueError(f"Invalid score: {value!r}") from error
    if not result.is_finite() or result < 0:
        raise ValueError(f"Score must be finite and nonnegative: {value!r}")
    return result


def validate(draft, profile):
    errors, warnings = [], []
    if draft.get("version") != 1:
        errors.append("draft.version must be 1")
    if draft.get("mode") not in ("existing-exam", "proposed-blueprint"):
        errors.append("mode must be existing-exam or proposed-blueprint")
    questions = draft.get("questions", [])
    topics = draft.get("topics", [])
    allocations = draft.get("allocations", [])
    if not questions or not topics or not allocations:
        errors.append("questions, topics and allocations must all be nonempty")
    capacity = profile["last_row"] - profile["first_row"] + 1
    if len(topics) > capacity:
        errors.append(f"{len(topics)} topics exceed template capacity {capacity}; consolidate topics or supply a larger template")
    for collection, label in ((questions, "question"), (topics, "topic"), (allocations, "allocation")):
        identifiers = [entry.get("id") for entry in collection]
        if any(not isinstance(x, str) or not x for x in identifiers) or len(identifiers) != len(set(identifiers)):
            errors.append(f"Every {label} needs a unique nonempty string id")
    question_map = {q["id"]: q for q in questions if isinstance(q.get("id"), str)}
    topic_map = {t["id"]: t for t in topics if isinstance(t.get("id"), str)}
    if any(not isinstance(t.get("title"), str) or not t["title"].strip() for t in topics):
        errors.append("Every topic needs a nonempty title")
    criterion_map = {}
    for question in questions:
        criteria = question.get("criteria", [])
        ids = [item.get("id") for item in criteria]
        if not criteria or any(not isinstance(x, str) or not x for x in ids) or len(ids) != len(set(ids)):
            errors.append(f"{question.get('id')}: criteria need unique nonempty ids")
        for criterion in criteria:
            criterion_map[(question.get("id"), criterion.get("id"))] = criterion
    topic_scores = {t: {g: Decimal(0) for g in GROUPS} for t in topic_map}
    question_scores = {q: {g: Decimal(0) for g in GROUPS} for q in question_map}
    criterion_totals = defaultdict(Decimal)
    subcriterion_totals = defaultdict(Decimal)
    seen_evidence = set()
    for entry in allocations:
        label = entry.get("id", "unnamed allocation")
        question, topic = entry.get("question"), entry.get("topic")
        key = (question, entry.get("criterion"))
        if question not in question_map or topic not in topic_map or key not in criterion_map:
            errors.append(f"{label}: unknown question, criterion or topic")
            continue
        if set(entry.get("scores", {})) != set(GROUPS):
            errors.append(f"{label}: scores must contain exactly {', '.join(GROUPS)}")
            continue
        try:
            scores = {g: number(entry["scores"][g]) for g in GROUPS}
            points = number(entry["points"])
            if sum(scores.values()) != points:
                errors.append(f"{label}: cognitive scores do not sum to points")
            for g in GROUPS:
                topic_scores[topic][g] += scores[g]
                question_scores[question][g] += scores[g]
            criterion_totals[key] += points
            declared = criterion_map[key].get("subcriteria", {})
            sub = entry.get("subcriterion")
            if (declared and sub not in declared) or (sub is not None and sub not in declared):
                errors.append(f"{label}: missing/unknown published subcriterion")
            if sub is not None:
                subcriterion_totals[(*key, sub)] += points
        except (ValueError, KeyError) as error:
            errors.append(f"{label}: {error}")
        if not entry.get("rationale", "").strip():
            errors.append(f"{label}: cognitive classification needs a rationale")
        evidence = entry.get("evidence", [])
        if not evidence or any(not e.get("source") or not e.get("locator") for e in evidence):
            errors.append(f"{label}: provide evidence with source and locator")
        for item in evidence:
            seen_evidence.add((item.get("source"), item.get("locator")))
    try:
        total = number(draft.get("total_points"))
        if not total:
            errors.append("total_points must be positive")
        if sum(number(q["points"]) for q in questions) != total:
            errors.append("Declared question totals do not equal total_points")
        for qid, question in question_map.items():
            if sum(question_scores[qid].values()) != number(question["points"]):
                errors.append(f"{qid}: allocations do not match declared question points")
            if sum(number(c["points"]) for c in question["criteria"]) != number(question["points"]):
                errors.append(f"{qid}: published criterion totals do not match question points")
            for criterion in question["criteria"]:
                key = (qid, criterion["id"])
                if criterion_totals[key] != number(criterion["points"]):
                    errors.append(f"{qid}-{criterion['id']}: allocation changes the criterion total")
                declared = criterion.get("subcriteria", {})
                if declared and sum(number(v) for v in declared.values()) != number(criterion["points"]):
                    errors.append(f"{qid}-{criterion['id']}: declared subcriteria do not sum to criterion points")
                for sub, points in declared.items():
                    if subcriterion_totals[(*key, sub)] != number(points):
                        errors.append(f"{qid}-{criterion['id']}/{sub}: allocation changes the published subcriterion total")
    except (ValueError, KeyError) as error:
        errors.append(str(error))
        total = Decimal(0)
    cognitive = {g: sum((s[g] for s in topic_scores.values()), Decimal(0)) for g in GROUPS}
    if sum(cognitive.values()) != total:
        errors.append("Allocated total does not equal total_points")
    shares = {g: cognitive[g] * 100 / total if total else Decimal(0) for g in GROUPS}
    band_results = {}
    for group, bounds in profile.get("bands", {}).items():
        low, high = map(number, bounds)
        band_results[group] = low <= shares[group] <= high
        if not band_results[group]:
            warnings.append(f"{group}: {shares[group]:.2f}% is outside {low}-{high}%. Do not relabel tasks to force compliance; propose assessment changes.")
    if not draft.get("metadata", {}).get("course_title"):
        warnings.append("Course title is missing; identify the field as unresolved.")
    input_mode = str(draft.get("metadata", {}).get("input_mode", "")).strip().lower()
    if input_mode in ("chat", "chat-pasted", "pasted"):
        warnings.append("input_mode is chat-pasted: published totals could not be checked against page images. Read every published criterion total back to the requester for confirmation before use.")
    warnings.extend(str(x) for x in draft.get("assumptions", []))
    return {"errors": errors, "warnings": warnings, "total": total,
            "cognitive": cognitive, "shares": shares, "topic_scores": topic_scores,
            "question_scores": question_scores, "bands": band_results,
            "evidence_locators": len(seen_evidence)}

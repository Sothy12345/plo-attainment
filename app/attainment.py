"""Attainment calculations built from real student scores.

The dashboards elsewhere in the app historically showed placeholder percentages.
Everything here is derived from `studentscore` rows:

    assessment scores -> CLO percent -> PLO attainment -> PEO attainment

A CLO percent is a student's total score on that CLO's assessments divided by
those assessments' combined max score, matching `services.get_clo_report`. Only
students who actually have a score row for the CLO are counted, so a programme
that is midway through score entry is not dragged down by unmarked students.
Programmes with no scores at all report `has_data = False` rather than 0%, which
would claim the students scored zero.
"""

from collections import defaultdict

from sqlmodel import Session, select

from app.models import (
    PEO,
    PLO,
    Assessment,
    CLO,
    CLOPLOMapping,
    Course,
    PEOPLOMapping,
    PLOTarget,
    Program,
    StudentScore,
)
from app.services import normalized_mapping_weight, plo_sort_key


DEFAULT_TARGET = 70.0


def _round(value: float | None) -> float | None:
    return None if value is None else round(value, 1)


def clo_percentages(
    session: Session, program_id: int, student_ids: set[int] | None = None
) -> dict[int, dict[int, float]]:
    """Return {clo_id: {student_id: percent 0-100}} for one programme.

    `student_ids` restricts the calculation to a cohort/year/semester selection;
    None means every student who has a score.
    """
    course_ids = [
        course.id
        for course in session.exec(select(Course).where(Course.program_id == program_id)).all()
        if course.id is not None
    ]
    if not course_ids:
        return {}
    clos = list(session.exec(select(CLO).where(CLO.course_id.in_(course_ids))).all())
    clo_ids = [clo.id for clo in clos if clo.id is not None]
    if not clo_ids:
        return {}

    assessments = list(session.exec(select(Assessment).where(Assessment.clo_id.in_(clo_ids))).all())
    assessment_clo = {item.id: item.clo_id for item in assessments if item.id is not None}
    clo_max: dict[int, float] = defaultdict(float)
    for item in assessments:
        clo_max[item.clo_id] += float(item.max_score or 0)

    assessment_ids = [item.id for item in assessments if item.id is not None]
    if not assessment_ids:
        return {}
    scores = list(
        session.exec(select(StudentScore).where(StudentScore.assessment_id.in_(assessment_ids))).all()
    )

    totals: dict[tuple[int, int], float] = defaultdict(float)
    for row in scores:
        clo_id = assessment_clo.get(row.assessment_id)
        if clo_id is None:
            continue
        if student_ids is not None and row.student_id not in student_ids:
            continue
        totals[(clo_id, row.student_id)] += float(row.score or 0)

    result: dict[int, dict[int, float]] = defaultdict(dict)
    for (clo_id, student_id), total in totals.items():
        maximum = clo_max.get(clo_id, 0)
        if maximum > 0:
            result[clo_id][student_id] = min(100.0, total / maximum * 100)
    return result


def plo_targets(
    session: Session,
    program_id: int,
    academic_year: str | None = None,
    cohort: str | None = None,
) -> dict[int, float]:
    """Configured target per PLO.

    A target matching the selected academic year and cohort wins; otherwise any
    configured target for the PLO is used. Callers fall back to the system
    default when a PLO has no target row at all.
    """
    exact: dict[int, float] = {}
    fallback: dict[int, float] = {}
    for row in session.exec(select(PLOTarget).where(PLOTarget.program_id == program_id)).all():
        if row.plo_id is None or not row.target:
            continue
        value = float(row.target)
        matches_year = not academic_year or str(row.academic_year or "") == academic_year
        matches_cohort = not cohort or str(row.cohort or "") == cohort
        if matches_year and matches_cohort:
            exact.setdefault(row.plo_id, value)
        fallback.setdefault(row.plo_id, value)
    return {**fallback, **exact}


def programme_plo_attainment(
    session: Session,
    program: Program,
    version_id: int | None,
    student_ids: set[int] | None = None,
    default_target: float = DEFAULT_TARGET,
    academic_year: str | None = None,
    cohort: str | None = None,
) -> dict:
    """Per-PLO attainment for one programme version, plus a programme summary."""
    plos = sorted(
        session.exec(
            select(PLO).where(PLO.program_id == program.id, PLO.plo_version_id == version_id)
        ).all(),
        key=plo_sort_key,
    )
    percents = clo_percentages(session, program.id, student_ids)
    targets = plo_targets(session, program.id, academic_year, cohort)

    plo_ids = [plo.id for plo in plos if plo.id is not None]
    mappings = (
        list(session.exec(select(CLOPLOMapping).where(CLOPLOMapping.plo_id.in_(plo_ids))).all())
        if plo_ids
        else []
    )
    by_plo: dict[int, list[CLOPLOMapping]] = defaultdict(list)
    for mapping in mappings:
        by_plo[mapping.plo_id].append(mapping)

    rows = []
    for plo in plos:
        # Weighted average of the student's mapped CLO percentages.
        student_weighted: dict[int, float] = defaultdict(float)
        student_weight: dict[int, float] = defaultdict(float)
        mapped_clos = 0
        for mapping in by_plo.get(plo.id, []):
            weight = normalized_mapping_weight(float(mapping.weight or 0))
            if weight <= 0:
                continue
            student_percents = percents.get(mapping.clo_id, {})
            if student_percents:
                mapped_clos += 1
            for student_id, percent in student_percents.items():
                student_weighted[student_id] += percent * weight
                student_weight[student_id] += weight

        student_scores = [
            student_weighted[student_id] / student_weight[student_id]
            for student_id in student_weighted
            if student_weight[student_id] > 0
        ]
        target = targets.get(plo.id, default_target)
        attainment = sum(student_scores) / len(student_scores) if student_scores else None
        met = sum(1 for score in student_scores if score >= target)
        rows.append(
            {
                "plo": plo,
                "code": plo.code,
                "description": plo.description,
                "attainment": _round(attainment),
                "target": round(target, 1),
                "students_assessed": len(student_scores),
                "students_meeting_target": met,
                "mapped_clos": mapped_clos,
                "has_data": attainment is not None,
                "status": _status(attainment, target),
            }
        )

    measured = [row["attainment"] for row in rows if row["attainment"] is not None]
    overall = sum(measured) / len(measured) if measured else None
    return {
        "program": program,
        "rows": rows,
        "overall": _round(overall),
        "has_data": overall is not None,
        "plo_count": len(rows),
        "measured_count": len(measured),
        "on_track": sum(1 for row in rows if row["status"] == "On Track"),
        "at_risk": sum(1 for row in rows if row["status"] == "At Risk"),
        "below_target": sum(1 for row in rows if row["status"] == "Below Target"),
    }


def programme_peo_attainment(
    session: Session,
    program: Program,
    version_id: int | None,
    plo_result: dict,
    default_target: float = DEFAULT_TARGET,
) -> dict:
    """Roll PLO attainment up to PEOs using the PEO-PLO contribution weights."""
    peos = list(
        session.exec(
            select(PEO)
            .where(PEO.program_id == program.id, PEO.plo_version_id == version_id)
            .order_by(PEO.code)
        ).all()
    )
    attainment_by_plo = {
        row["plo"].id: row["attainment"] for row in plo_result["rows"] if row["attainment"] is not None
    }
    links = list(
        session.exec(select(PEOPLOMapping).where(PEOPLOMapping.plo_version_id == version_id)).all()
    )
    by_peo: dict[int, list[PEOPLOMapping]] = defaultdict(list)
    for link in links:
        if link.is_mapped:
            by_peo[link.peo_id].append(link)

    rows = []
    for peo in peos:
        weighted = 0.0
        total_weight = 0.0
        contributing = 0
        for link in by_peo.get(peo.id, []):
            attainment = attainment_by_plo.get(link.plo_id)
            if attainment is None:
                continue
            weight = float(link.contribution_percentage or 0) or float(link.weight or 0)
            if weight <= 0:
                continue
            weighted += attainment * weight
            total_weight += weight
            contributing += 1
        value = weighted / total_weight if total_weight > 0 else None
        rows.append(
            {
                "peo": peo,
                "code": peo.code,
                "description": peo.description,
                "attainment": _round(value),
                "target": round(default_target, 1),
                "contributing_plos": contributing,
                "mapped_plos": len(by_peo.get(peo.id, [])),
                "has_data": value is not None,
                "status": _status(value, default_target),
            }
        )

    measured = [row["attainment"] for row in rows if row["attainment"] is not None]
    overall = sum(measured) / len(measured) if measured else None
    return {
        "program": program,
        "rows": rows,
        "overall": _round(overall),
        "has_data": overall is not None,
        "peo_count": len(rows),
        "measured_count": len(measured),
        "achieved": sum(1 for row in rows if row["status"] == "On Track"),
    }


def _status(value: float | None, target: float) -> str:
    if value is None:
        return "No Data"
    if value >= target:
        return "On Track"
    if value >= target - 10:
        return "At Risk"
    return "Below Target"


def faculty_attainment(
    session: Session,
    programs: list[Program],
    version_for,
    students_for=None,
    default_target: float = DEFAULT_TARGET,
    academic_year: str | None = None,
    cohort: str | None = None,
) -> dict:
    """Per-programme PLO and PEO attainment plus the faculty averages.

    `version_for` resolves the outcome version a programme should be reported on,
    so callers keep control of version selection.
    """
    programmes = []
    for program in programs:
        version = version_for(program)
        version_id = version.id if version else None
        student_ids = students_for(program) if students_for else None
        plo_result = programme_plo_attainment(
            session, program, version_id, student_ids, default_target, academic_year, cohort
        )
        peo_result = programme_peo_attainment(
            session, program, version_id, plo_result, default_target
        )
        programmes.append(
            {
                "program": program,
                "code": program.code,
                "name": program.name,
                "version": version,
                "plo": plo_result,
                "peo": peo_result,
            }
        )

    plo_values = [item["plo"]["overall"] for item in programmes if item["plo"]["overall"] is not None]
    peo_values = [item["peo"]["overall"] for item in programmes if item["peo"]["overall"] is not None]
    return {
        "programmes": programmes,
        "programme_count": len(programmes),
        "plo_average": _round(sum(plo_values) / len(plo_values)) if plo_values else None,
        "peo_average": _round(sum(peo_values) / len(peo_values)) if peo_values else None,
        "plo_reporting_count": len(plo_values),
        "peo_reporting_count": len(peo_values),
        "plo_total": sum(item["plo"]["plo_count"] for item in programmes),
        "peo_total": sum(item["peo"]["peo_count"] for item in programmes),
        "on_track": sum(item["plo"]["on_track"] for item in programmes),
        "at_risk": sum(item["plo"]["at_risk"] for item in programmes),
        "below_target": sum(item["plo"]["below_target"] for item in programmes),
    }


# --- Continuous Quality Improvement -------------------------------------------------

CQI_ACTIONS = {
    "curriculum": "Curriculum review",
    "teaching": "Teaching and learning improvement",
    "assessment": "Assessment review",
    "support": "Student support",
    "mapping": "CLO-PLO mapping review",
    "peo_mapping": "PEO-PLO mapping review",
    "data": "Complete score entry",
}


def plo_cqi_actions(row: dict) -> list[dict]:
    """Suggest CQI actions for one PLO from its own measured signals.

    Every trigger below is read off the row, so the reason shown to the Dean is
    always traceable to real numbers rather than a canned list.
    """
    actions: list[dict] = []
    attainment = row.get("attainment")
    target = row.get("target") or DEFAULT_TARGET
    assessed = row.get("students_assessed") or 0
    met = row.get("students_meeting_target") or 0
    mapped = row.get("mapped_clos") or 0

    if attainment is None:
        if mapped == 0:
            actions.append({"action": CQI_ACTIONS["mapping"],
                            "reason": "No CLO is mapped to this PLO, so nothing can be measured."})
        else:
            actions.append({"action": CQI_ACTIONS["data"],
                            "reason": "Mapped CLOs exist but no marks have been entered yet."})
        return actions

    gap = target - attainment
    if gap <= 0:
        return actions

    if gap >= 15:
        actions.append({"action": CQI_ACTIONS["curriculum"],
                        "reason": f"Attainment is {round(gap, 1)}% below the {target}% target."})
    else:
        actions.append({"action": CQI_ACTIONS["teaching"],
                        "reason": f"Attainment is {round(gap, 1)}% below the {target}% target."})

    if mapped and mapped < 2:
        actions.append({"action": CQI_ACTIONS["mapping"],
                        "reason": f"Only {mapped} CLO feeds this PLO, so the result rests on a single measure."})

    if assessed:
        met_share = met / assessed * 100
        if met_share < 50:
            actions.append({"action": CQI_ACTIONS["support"],
                            "reason": f"Only {met} of {assessed} students reached the target ({round(met_share, 1)}%)."})
        elif met_share >= 70 and attainment < target:
            actions.append({"action": CQI_ACTIONS["assessment"],
                            "reason": f"{met} of {assessed} students passed the target yet the average is below it, which suggests uneven assessment weighting."})
    return actions


def peo_cqi_actions(row: dict) -> list[dict]:
    """Suggest CQI actions for one PEO from its PLO contributions."""
    actions: list[dict] = []
    attainment = row.get("attainment")
    target = row.get("target") or DEFAULT_TARGET
    contributing = row.get("contributing_plos") or 0
    mapped = row.get("mapped_plos") or 0

    if attainment is None:
        if mapped == 0:
            actions.append({"action": CQI_ACTIONS["peo_mapping"],
                            "reason": "No PLO is mapped to this PEO, so it cannot be measured."})
        else:
            actions.append({"action": CQI_ACTIONS["data"],
                            "reason": f"{mapped} PLO(s) are mapped but none of them has attainment data yet."})
        return actions

    gap = target - attainment
    if gap > 0:
        if gap >= 15:
            actions.append({"action": CQI_ACTIONS["curriculum"],
                            "reason": f"Attainment is {round(gap, 1)}% below the {target}% target."})
        else:
            actions.append({"action": CQI_ACTIONS["teaching"],
                            "reason": f"Attainment is {round(gap, 1)}% below the {target}% target."})

    if mapped and contributing < mapped:
        actions.append({"action": CQI_ACTIONS["peo_mapping"],
                        "reason": f"Only {contributing} of {mapped} mapped PLO(s) carry data, so this PEO is measured on a partial mapping."})
    return actions


def cqi_report(programmes: list[dict], focus: str) -> list[dict]:
    """Flatten per-programme outcomes into a Dean-facing CQI action list."""
    report = []
    for entry in programmes:
        result = entry["peo"] if focus == "peo" else entry["plo"]
        for row in result["rows"]:
            actions = peo_cqi_actions(row) if focus == "peo" else plo_cqi_actions(row)
            if not actions:
                continue
            report.append(
                {
                    "program": entry["program"],
                    "programme_code": entry["code"],
                    "programme_name": entry["name"],
                    "code": row["code"],
                    "description": row["description"],
                    "attainment": row["attainment"],
                    "target": row["target"],
                    # Binary status, matching the Dean comparison table.
                    "status": (
                        "No Data" if row["attainment"] is None
                        else ("Achieved" if row["attainment"] >= (row["target"] or DEFAULT_TARGET) else "Below Target")
                    ),
                    "actions": actions,
                    "severity": (
                        0 if row["attainment"] is None else round((row["target"] or DEFAULT_TARGET) - row["attainment"], 1)
                    ),
                }
            )
    report.sort(key=lambda item: (-item["severity"], item["programme_code"], item["code"]))
    return report

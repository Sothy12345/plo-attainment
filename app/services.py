from collections import defaultdict

from sqlmodel import Session, select

from app.models import Assessment, CLO, CLOPLOMapping, ClassStudent, CourseClass, PLO, Program, Student, StudentScore


def plo_sort_key(plo: PLO) -> int:
    digits = "".join(character for character in plo.code if character.isdigit())
    return int(digits or 0)


def normalized_mapping_weight(weight: float) -> float:
    if weight > 1:
        return weight / 100
    return weight


def get_class_students(session: Session, class_id: int) -> list[Student]:
    enrollments = session.exec(select(ClassStudent).where(ClassStudent.class_id == class_id)).all()
    return [enrollment.student for enrollment in enrollments]


def get_clo_report(session: Session, class_id: int, clo_id: int) -> dict:
    course_class = session.get(CourseClass, class_id)
    clo = session.get(CLO, clo_id)
    assessments = list(session.exec(select(Assessment).where(Assessment.clo_id == clo_id)))
    students = [enrollment.student for enrollment in course_class.students]
    max_score = sum(item.max_score for item in assessments)
    threshold = clo.pass_threshold if clo else 0.5

    score_rows = list(
        session.exec(
            select(StudentScore)
            .join(Assessment)
            .where(Assessment.clo_id == clo_id)
        )
    )
    score_map = {(row.student_id, row.assessment_id): row.score for row in score_rows}

    rows = []
    pass_count = 0
    for student in students:
        marks = []
        total = 0.0
        for assessment in assessments:
            score = score_map.get((student.id, assessment.id), 0)
            total += score
            marks.append({"assessment": assessment, "score": score})

        percent = total / max_score if max_score else 0
        status = "Pass" if percent >= threshold else "Fail"
        if status == "Pass":
            pass_count += 1
        rows.append(
            {
                "student": student,
                "marks": marks,
                "total": total,
                "percent": percent,
                "status": status,
            }
        )

    total_students = len(rows)
    pass_percent = pass_count / total_students if total_students else 0
    return {
        "class": course_class,
        "clo": clo,
        "assessments": assessments,
        "rows": rows,
        "max_score": max_score,
        "pass_count": pass_count,
        "fail_count": total_students - pass_count,
        "pass_percent": pass_percent,
        "achievement": "Yes" if pass_percent >= threshold else "No",
    }


def get_course_report(session: Session, class_id: int) -> dict:
    course_class = session.get(CourseClass, class_id)
    clos = list(session.exec(select(CLO).where(CLO.course_id == course_class.course_id)))
    clo_reports = [get_clo_report(session, class_id, clo.id) for clo in clos]

    students = [enrollment.student for enrollment in course_class.students]
    student_totals: dict[int, float] = defaultdict(float)
    student_max: dict[int, float] = defaultdict(float)
    for report in clo_reports:
        for row in report["rows"]:
            student_totals[row["student"].id] += row["total"]
            student_max[row["student"].id] += report["max_score"]

    rows = []
    pass_count = 0
    for student in students:
        total = student_totals[student.id]
        max_score = student_max[student.id]
        percent = total / max_score if max_score else 0
        clo_marks = []
        for report in clo_reports:
            clo_row = next((row for row in report["rows"] if row["student"].id == student.id), None)
            clo_marks.append(
                {
                    "clo": report["clo"],
                    "score": clo_row["total"] if clo_row else 0,
                    "percent": clo_row["percent"] if clo_row else 0,
                    "status": clo_row["status"] if clo_row else "Fail",
                }
            )
        status = "Pass" if percent >= 0.5 else "Fail"
        if status == "Pass":
            pass_count += 1
        rows.append(
            {
                "student": student,
                "clo_marks": clo_marks,
                "total": total,
                "percent": percent,
                "status": status,
                "grade": grade_from_percent(percent),
                "cqi": "",
            }
        )

    total_students = len(rows)
    pass_percent = pass_count / total_students if total_students else 0
    fail_count = total_students - pass_count
    return {
        "class": course_class,
        "clos": clos,
        "clo_reports": clo_reports,
        "rows": rows,
        "pass_count": pass_count,
        "fail_count": fail_count,
        "pass_percent": pass_percent,
        "fail_percent": fail_count / total_students if total_students else 0,
        "course_condition": 0.5,
        "achievement": "Yes" if pass_percent >= 0.5 else "No",
    }


def get_plo_summary(session: Session, class_id: int) -> list[dict]:
    course_report = get_course_report(session, class_id)
    summary = []
    for clo_report in course_report["clo_reports"]:
        mappings = list(session.exec(select(CLOPLOMapping).where(CLOPLOMapping.clo_id == clo_report["clo"].id)))
        for mapping in mappings:
            summary.append(
                {
                    "plo": mapping.plo,
                    "clo": clo_report["clo"],
                    "weight": mapping.weight,
                    "weight_ratio": normalized_mapping_weight(mapping.weight),
                    "pass_percent": clo_report["pass_percent"],
                    "achievement": clo_report["achievement"],
                }
            )
    return summary


def get_program_report(session: Session, program_id: int) -> dict:
    program = session.get(Program, program_id)
    plos = sorted(program.plos, key=plo_sort_key)
    # SQLModel relationship joins are intentionally avoided here for compatibility.
    classes = [course_class for course_class in session.exec(select(CourseClass)).all() if course_class.course.program_id == program_id]

    student_rows: dict[int, dict] = {}
    plo_totals: dict[int, float] = defaultdict(float)
    plo_counts: dict[int, int] = defaultdict(int)

    for course_class in classes:
        course_report = get_course_report(session, course_class.id)
        for row in course_report["rows"]:
            student = row["student"]
            if student.id not in student_rows:
                student_rows[student.id] = {
                    "student": student,
                    "plo_entries": {plo.id: [] for plo in plos},
                }

            for clo_mark in row["clo_marks"]:
                mappings = session.exec(select(CLOPLOMapping).where(CLOPLOMapping.clo_id == clo_mark["clo"].id)).all()
                for mapping in mappings:
                    value = clo_mark["percent"] * 4 * normalized_mapping_weight(mapping.weight)
                    student_rows[student.id]["plo_entries"].setdefault(mapping.plo_id, []).append(value)

    rows = []
    for item in student_rows.values():
        values = []
        for plo in plos:
            entries = item["plo_entries"].get(plo.id, [])
            value = sum(entries) / len(entries) if entries else 0
            values.append(value)
            plo_totals[plo.id] += value
            plo_counts[plo.id] += 1
        rows.append({"student": item["student"], "plo_values": values})

    averages = [
        (plo_totals[plo.id] / plo_counts[plo.id]) if plo_counts[plo.id] else 0
        for plo in plos
    ]
    return {"program": program, "plos": plos, "classes": classes, "rows": rows, "averages": averages}


def grade_from_percent(percent: float) -> str:
    if percent >= 0.85:
        return "A"
    if percent >= 0.80:
        return "B+"
    if percent >= 0.70:
        return "B"
    if percent >= 0.65:
        return "C+"
    if percent >= 0.50:
        return "C"
    if percent >= 0.45:
        return "D"
    if percent >= 0.40:
        return "E"
    return "F"

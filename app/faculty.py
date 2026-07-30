"""Real faculty-scoped data for the Dean pages.

The Dean screens used to render hardcoded figures (a fixed 48 faculty members,
"78.4%" attainment, invented programme names). Everything here is queried from
the database instead, so a Dean sees their own faculty.

Attainment values come from `app.attainment`, which derives them from entered
student scores. Where a programme has no scores the value is `None` and the UI
shows "No data" rather than 0%.
"""

from collections import defaultdict

from sqlmodel import Session, select

from app.attainment import faculty_attainment
from app.models import (
    PEO,
    PLO,
    Assessment,
    CLO,
    ClassStudent,
    Course,
    CourseClass,
    CourseTeacher,
    Faculty,
    Program,
    Role,
    StudentScore,
    StudentSemesterEnrollment,
    Teacher,
    User,
)


ROLE_TITLES = {
    Role.DEAN: "Dean",
    Role.PROGRAM_MANAGER: "Programme Coordinator",
    Role.TEACHER: "Lecturer",
}


def faculty_programmes(session: Session, faculty_id: int | None) -> list[Program]:
    if faculty_id is None:
        return []
    return list(
        session.exec(
            select(Program).where(Program.faculty_id == faculty_id).order_by(Program.code)
        ).all()
    )


def _programme_ids(programmes: list[Program]) -> list[int]:
    return [item.id for item in programmes if item.id is not None]


def faculty_members(session: Session, faculty_id: int | None, programmes: list[Program]) -> list[dict]:
    """Staff attached to the faculty, either directly or through a programme."""
    programme_ids = set(_programme_ids(programmes))
    programme_by_id = {item.id: item for item in programmes}
    staff_roles = {Role.DEAN, Role.PROGRAM_MANAGER, Role.TEACHER}
    users = [
        item
        for item in session.exec(select(User).order_by(User.role, User.name)).all()
        if item.role in staff_roles
        and (item.program_id in programme_ids or (faculty_id is not None and item.faculty_id == faculty_id))
    ]
    user_ids = [item.id for item in users if item.id is not None]
    teachers = (
        list(session.exec(select(Teacher).where(Teacher.user_id.in_(user_ids))).all()) if user_ids else []
    )
    teacher_by_user = {item.user_id: item for item in teachers}
    teacher_ids = [item.id for item in teachers if item.id is not None]
    course_counts: dict[int, int] = defaultdict(int)
    if teacher_ids:
        for row in session.exec(
            select(CourseTeacher).where(CourseTeacher.teacher_id.in_(teacher_ids))
        ).all():
            course_counts[row.teacher_id] += 1

    rows = []
    for item in users:
        teacher = teacher_by_user.get(item.id)
        programme = programme_by_id.get(item.program_id)
        rows.append(
            {
                "user": item,
                "name": item.name,
                "email": item.email,
                "role": item.role,
                "title": ROLE_TITLES.get(item.role, "Staff"),
                "staff_no": teacher.staff_no if teacher else "-",
                "programme": programme.code if programme else "Faculty-wide",
                "programme_name": programme.name if programme else "",
                "courses": course_counts.get(teacher.id, 0) if teacher else 0,
                "is_active": item.is_active,
            }
        )
    rows.sort(key=lambda row: (row["role"] != Role.DEAN, row["role"] != Role.PROGRAM_MANAGER, row["name"]))
    return rows


def programme_rows(session: Session, programmes: list[Program], summary: dict) -> list[dict]:
    """One row per programme: coordinator, size and real PLO attainment."""
    programme_ids = _programme_ids(programmes)
    coordinators: dict[int, str] = {}
    for item in session.exec(select(User).where(User.role == Role.PROGRAM_MANAGER)).all():
        if item.program_id in programme_ids and item.program_id not in coordinators:
            coordinators[item.program_id] = item.name

    student_counts: dict[int, int] = defaultdict(int)
    seen: set[tuple[int, int]] = set()
    for row in session.exec(select(StudentSemesterEnrollment)).all():
        if row.program_id in programme_ids and (row.program_id, row.student_id) not in seen:
            seen.add((row.program_id, row.student_id))
            student_counts[row.program_id] += 1

    course_counts: dict[int, int] = defaultdict(int)
    for course in session.exec(select(Course)).all():
        if course.program_id in programme_ids:
            course_counts[course.program_id] += 1

    attainment_by_programme = {item["program"].id: item for item in summary["programmes"]}

    rows = []
    for programme in programmes:
        entry = attainment_by_programme.get(programme.id, {})
        plo_result = entry.get("plo", {})
        peo_result = entry.get("peo", {})
        rows.append(
            {
                "program": programme,
                "code": programme.code,
                "name": programme.name,
                "level": "Bachelor",
                "coordinator": coordinators.get(programme.id, "Not assigned"),
                "students": student_counts.get(programme.id, 0),
                "courses": course_counts.get(programme.id, 0),
                "plo_count": plo_result.get("plo_count", 0),
                "peo_count": peo_result.get("peo_count", 0),
                "version": entry.get("version"),
                "attainment": plo_result.get("overall"),
                "peo_attainment": peo_result.get("overall"),
                "status": _status_from(plo_result.get("overall")),
            }
        )
    return rows


def _status_from(value: float | None, target: float = 70.0) -> str:
    if value is None:
        return "No Data"
    if value >= target:
        return "On Track"
    if value >= target - 10:
        return "At Risk"
    return "Below Target"


def assessment_rows(session: Session, programmes: list[Program]) -> list[dict]:
    """Every assessment in the faculty with its real score coverage."""
    programme_ids = _programme_ids(programmes)
    programme_by_id = {item.id: item for item in programmes}
    courses = [item for item in session.exec(select(Course)).all() if item.program_id in programme_ids]
    course_by_id = {item.id: item for item in courses}
    course_ids = [item.id for item in courses if item.id is not None]
    if not course_ids:
        return []
    clos = list(session.exec(select(CLO).where(CLO.course_id.in_(course_ids))).all())
    clo_by_id = {item.id: item for item in clos}
    clo_ids = [item.id for item in clos if item.id is not None]
    if not clo_ids:
        return []
    assessments = list(session.exec(select(Assessment).where(Assessment.clo_id.in_(clo_ids))).all())
    assessment_ids = [item.id for item in assessments if item.id is not None]
    scores = (
        list(session.exec(select(StudentScore).where(StudentScore.assessment_id.in_(assessment_ids))).all())
        if assessment_ids
        else []
    )
    totals: dict[int, list[float]] = defaultdict(list)
    for row in scores:
        totals[row.assessment_id].append(float(row.score or 0))

    rows = []
    for item in assessments:
        clo = clo_by_id.get(item.clo_id)
        course = course_by_id.get(clo.course_id) if clo else None
        programme = programme_by_id.get(course.program_id) if course else None
        marks = totals.get(item.id, [])
        maximum = float(item.max_score or 0)
        average = (sum(marks) / len(marks) / maximum * 100) if marks and maximum > 0 else None
        rows.append(
            {
                "assessment": item,
                "name": item.name,
                "clo": clo.code if clo else "-",
                "course": course.code if course else "-",
                "course_title": course.title if course else "",
                "programme": programme.code if programme else "-",
                "max_score": maximum,
                "scored": len(marks),
                "attainment": round(average, 1) if average is not None else None,
                "status": "Completed" if marks else "Not Started",
            }
        )
    rows.sort(key=lambda row: (row["programme"], row["course"], row["name"]))
    return rows


def faculty_overview(session: Session, user: User, version_for) -> dict:
    """Everything the Dean pages need, computed once from the database."""
    faculty = session.get(Faculty, user.faculty_id) if user.faculty_id else None
    programmes = faculty_programmes(session, user.faculty_id)
    programme_ids = _programme_ids(programmes)
    summary = faculty_attainment(session, programmes, version_for)
    members = faculty_members(session, user.faculty_id, programmes)
    rows = programme_rows(session, programmes, summary)
    assessments = assessment_rows(session, programmes)

    classes = [
        item
        for item in session.exec(select(CourseClass)).all()
        if item.course and item.course.program_id in programme_ids
    ]
    class_ids = {item.id for item in classes if item.id is not None}
    enrolled = {
        row.student_id
        for row in session.exec(select(ClassStudent)).all()
        if row.class_id in class_ids
    }
    plo_total = sum(
        1 for item in session.exec(select(PLO)).all() if item.program_id in programme_ids
    )
    peo_total = sum(
        1 for item in session.exec(select(PEO)).all() if item.program_id in programme_ids
    )
    scored = sum(1 for row in assessments if row["scored"])

    return {
        "faculty": faculty,
        "programmes": programmes,
        "summary": summary,
        "members": members,
        "programme_rows": rows,
        "assessments": assessments,
        "counts": {
            "programmes": len(programmes),
            "members": len(members),
            "lecturers": sum(1 for item in members if item["role"] == Role.TEACHER),
            "coordinators": sum(1 for item in members if item["role"] == Role.PROGRAM_MANAGER),
            "students": len(enrolled) or sum(row["students"] for row in rows),
            "courses": sum(row["courses"] for row in rows),
            "classes": len(classes),
            "assessments": len(assessments),
            "assessments_scored": scored,
            "assessments_pending": len(assessments) - scored,
            "plos": plo_total,
            "peos": peo_total,
        },
    }

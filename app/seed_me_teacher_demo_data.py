"""Seed Mechanical Engineering demo data for testing teacher score input and promotion.

Run from project root:
    python seed_me_teacher_demo_data.py

Safe behavior:
- Does not delete existing data.
- Updates/creates demo users, ME courses, 8 class codes, teacher assignments,
  students, assessments, sample scores, and audit logs.
"""
from __future__ import annotations

from datetime import datetime
from random import Random

from sqlmodel import Session, select

from app.database import create_db_and_tables, engine
from app.models import (
    AcademicYear,
    Assessment,
    AuditLog,
    CLO,
    CLOPLOMapping,
    ClassStudent,
    ClassTeacher,
    Course,
    CourseClass,
    CourseTeacher,
    Faculty,
    PLO,
    Program,
    Role,
    Student,
    StudentScore,
    StudentSemesterEnrollment,
    Teacher,
    User,
)
from app.security import hash_password
from app.seed import ME_CURRICULUM, ME_PLOS, sync_faculty_program_structure

PASSWORD = "password"
GENERATION = "21"
PROGRAM_CODE = "ME"
SHIFT_CODE = "M"  # Morning
DEGREE_CODE = "b"  # Bachelor
GROUP_CODE = "1"

YEAR_TO_ACADEMIC_YEAR = {
    1: "2025-2026",
    2: "2026-2027",
    3: "2027-2028",
    4: "2028-2029",
}

TEACHERS = [
    ("ME-T001", "Teacher Sok Dara", "me.teacher1@example.com"),
    ("ME-T002", "Teacher Chan Vireak", "me.teacher2@example.com"),
    ("ME-T003", "Teacher Long Sophea", "me.teacher3@example.com"),
    ("ME-T004", "Teacher Kim Sopheak", "me.teacher4@example.com"),
    ("ME-T005", "Teacher Heng Piseth", "me.teacher5@example.com"),
    ("ME-T006", "Teacher Neang Sovan", "me.teacher6@example.com"),
]

KH_LAST = ["សុខ", "ចាន់", "លី", "ហេង", "គង់", "ស៊ុន", "វ៉ាន់", "ផាន", "ឃុន", "ញ៉ែម"]
KH_FIRST = ["ដារ៉ា", "វីរៈ", "សុភា", "បញ្ញា", "សុធី", "វិចិត្រ", "ពិសិដ្ឋ", "ចំរើន", "សំណាង", "ធីតា"]
EN_NAMES = [
    "Sok Dara", "Chan Vireak", "Ly Sophea", "Heng Panha", "Kong Sothea",
    "Sun Vichetra", "Van Piseth", "Phan Chamroeun", "Khun Samnang", "Nyem Thida",
]


def class_code(year: int, semester: int) -> str:
    return f"{GENERATION}{PROGRAM_CODE}{year}{semester}{SHIFT_CODE}{DEGREE_CODE}{GROUP_CODE}"


def get_or_create_user(session: Session, *, name: str, email: str, role: Role, faculty_id: int | None, program_id: int | None) -> User:
    user = session.exec(select(User).where(User.email == email)).first()
    if user:
        user.name = name
        user.role = role
        user.is_active = True
        user.faculty_id = faculty_id
        user.program_id = program_id
        session.add(user)
        session.commit()
        session.refresh(user)
        return user
    user = User(
        name=name,
        email=email,
        password_hash=hash_password(PASSWORD),
        role=role,
        faculty_id=faculty_id,
        program_id=program_id,
        is_active=True,
    )
    session.add(user)
    session.commit()
    session.refresh(user)
    return user


def get_or_create_teacher(session: Session, user: User, staff_no: str) -> Teacher:
    teacher = session.exec(select(Teacher).where(Teacher.user_id == user.id)).first()
    if teacher:
        teacher.staff_no = staff_no
        session.add(teacher)
        session.commit()
        session.refresh(teacher)
        return teacher
    teacher = Teacher(user_id=user.id, staff_no=staff_no)
    session.add(teacher)
    session.commit()
    session.refresh(teacher)
    return teacher


def ensure_academic_years(session: Session) -> None:
    for name in YEAR_TO_ACADEMIC_YEAR.values():
        existing = session.exec(select(AcademicYear).where(AcademicYear.name == name)).first()
        if not existing:
            start = name.split("-")[0]
            end = name.split("-")[1]
            session.add(AcademicYear(name=name, start_date=f"{start}-08-01", end_date=f"{end}-07-31", is_active=True, is_default=(name == "2026-2027")))
        elif name == "2026-2027":
            existing.is_default = True
            existing.is_active = True
            session.add(existing)
    session.commit()


def ensure_me_plos(session: Session, program: Program) -> list[PLO]:
    plos = []
    for code, description in ME_PLOS:
        plo = session.exec(select(PLO).where(PLO.program_id == program.id, PLO.code == code)).first()
        if not plo:
            plo = PLO(program_id=program.id, code=code, description=description)
            session.add(plo)
            session.commit()
            session.refresh(plo)
        plos.append(plo)
    return plos


def ensure_courses(session: Session, program: Program) -> list[Course]:
    courses = []
    for year, semester, code, title, credits in ME_CURRICULUM:
        course = session.exec(
            select(Course).where(
                Course.program_id == program.id,
                Course.code == code,
                Course.curriculum_year == year,
                Course.curriculum_semester == str(semester),
            )
        ).first()
        if not course:
            course = Course(program_id=program.id, code=code, title=title, credits=float(credits), curriculum_year=year, curriculum_semester=str(semester))
            session.add(course)
        else:
            course.title = title
            course.credits = float(credits)
            course.curriculum_year = year
            course.curriculum_semester = str(semester)
            session.add(course)
        session.commit()
        session.refresh(course)
        courses.append(course)
    return courses


def ensure_assessment_plan(session: Session, course: Course, plos: list[PLO]) -> list[Assessment]:
    clos = session.exec(select(CLO).where(CLO.course_id == course.id).order_by(CLO.code)).all()
    if not clos:
        clos = [
            CLO(course_id=course.id, code="CLO1", domain="K,S", description=f"Explain and apply key concepts in {course.title}."),
            CLO(course_id=course.id, code="CLO2", domain="S,A", description=f"Solve practical tasks and demonstrate professional responsibility in {course.title}."),
        ]
        session.add_all(clos)
        session.commit()
        for clo in clos:
            session.refresh(clo)
    for idx, clo in enumerate(clos[:2]):
        target_plo = plos[(int(course.curriculum_year or 1) + idx) % len(plos)]
        exists = session.exec(select(CLOPLOMapping).where(CLOPLOMapping.clo_id == clo.id, CLOPLOMapping.plo_id == target_plo.id)).first()
        if not exists:
            session.add(CLOPLOMapping(clo_id=clo.id, plo_id=target_plo.id, weight=50))
    session.commit()

    assessment_plan = [
        (clos[0], "Attendance", 10, 10),
        (clos[0], "Assignment", 20, 20),
        (clos[min(1, len(clos)-1)], "Midterm", 30, 30),
        (clos[min(1, len(clos)-1)], "Final Exam", 40, 40),
    ]
    assessments: list[Assessment] = []
    for clo, name, max_score, weight in assessment_plan:
        assessment = session.exec(select(Assessment).where(Assessment.clo_id == clo.id, Assessment.name == name)).first()
        if not assessment:
            assessment = Assessment(clo_id=clo.id, name=name, description=f"{name} for {course.code}", max_score=max_score, weight=weight)
            session.add(assessment)
        else:
            assessment.max_score = max_score
            assessment.weight = weight
            assessment.description = f"{name} for {course.code}"
            session.add(assessment)
        session.commit()
        session.refresh(assessment)
        assessments.append(assessment)
    return assessments


def ensure_assignment(session: Session, course: Course, course_class: CourseClass, teacher: Teacher) -> None:
    if not session.exec(select(CourseTeacher).where(CourseTeacher.course_id == course.id, CourseTeacher.teacher_id == teacher.id)).first():
        session.add(CourseTeacher(course_id=course.id, teacher_id=teacher.id))
    if not session.exec(select(ClassTeacher).where(ClassTeacher.class_id == course_class.id, ClassTeacher.teacher_id == teacher.id)).first():
        session.add(ClassTeacher(class_id=course_class.id, teacher_id=teacher.id))
    session.commit()


def ensure_students_for_year(session: Session, program: Program, faculty: Faculty, year: int) -> list[Student]:
    students: list[Student] = []
    for i in range(1, 11):
        student_no = f"{GENERATION}{PROGRAM_CODE}{year}{i:02d}"
        name_en = f"ME Year {year} {EN_NAMES[i-1]}"
        name_kh = f"{KH_LAST[i-1]} {KH_FIRST[i-1]}"
        email = f"student.{student_no.lower()}@example.com"
        user = session.exec(select(User).where(User.email == email)).first()
        if not user:
            user = User(name=name_en, email=email, password_hash=hash_password(PASSWORD), role=Role.STUDENT, faculty_id=faculty.id, program_id=program.id, is_active=True)
            session.add(user)
            session.commit()
            session.refresh(user)
        student = session.exec(select(Student).where(Student.student_no == student_no)).first()
        if not student:
            student = Student(user_id=user.id, student_no=student_no, name_kh=name_kh, name_en=name_en)
            session.add(student)
        else:
            student.user_id = user.id
            student.name_kh = name_kh
            student.name_en = name_en
            session.add(student)
        session.commit()
        session.refresh(student)
        students.append(student)
    return students


def ensure_class_student(session: Session, course_class: CourseClass, student: Student) -> None:
    if not session.exec(select(ClassStudent).where(ClassStudent.class_id == course_class.id, ClassStudent.student_id == student.id)).first():
        session.add(ClassStudent(class_id=course_class.id, student_id=student.id, status="Active"))
        session.commit()


def ensure_semester_enrollment(session: Session, program: Program, student: Student, code: str, academic_year: str, semester_label: str) -> None:
    exists = session.exec(
        select(StudentSemesterEnrollment).where(
            StudentSemesterEnrollment.student_id == student.id,
            StudentSemesterEnrollment.program_id == program.id,
            StudentSemesterEnrollment.cohort_name == code,
            StudentSemesterEnrollment.academic_year == academic_year,
            StudentSemesterEnrollment.semester == semester_label,
        )
    ).first()
    if not exists:
        session.add(StudentSemesterEnrollment(student_id=student.id, program_id=program.id, cohort_name=code, academic_year=academic_year, semester=semester_label, status="Active"))
        session.commit()


def seed_scores_for_teacher(session: Session, students: list[Student], assessments: list[Assessment], teacher_user_id: int, seed_value: int) -> None:
    rng = Random(seed_value)
    now = datetime.utcnow()
    for student in students:
        for assessment in assessments:
            max_score = float(assessment.max_score or 100)
            base = 0.68 + rng.random() * 0.25
            score_value = round(max_score * base, 2)
            score = session.exec(select(StudentScore).where(StudentScore.assessment_id == assessment.id, StudentScore.student_id == student.id)).first()
            if score:
                if score.locked:
                    continue
                score.score = score_value
                score.status = "Draft"
                score.updated_at = now
                score.entered_by_user_id = teacher_user_id
                session.add(score)
            else:
                session.add(StudentScore(assessment_id=assessment.id, student_id=student.id, score=score_value, status="Draft", updated_at=now, entered_by_user_id=teacher_user_id))
    session.commit()


def run() -> None:
    create_db_and_tables()
    with Session(engine) as session:
        sync_faculty_program_structure(session)
        ensure_academic_years(session)
        faculty = session.exec(select(Faculty).where(Faculty.name == "Faculty of Science and Technology")).first()
        program = session.exec(select(Program).where(Program.code == PROGRAM_CODE)).first()
        if not faculty or not program:
            raise RuntimeError("Faculty of Science and Technology or Mechanical Engineering programme not found.")

        manager = get_or_create_user(session, name="ME Programme Coordinator", email="me.manager@example.com", role=Role.PROGRAM_MANAGER, faculty_id=faculty.id, program_id=program.id)
        teachers: list[Teacher] = []
        for staff_no, name, email in TEACHERS:
            user = get_or_create_user(session, name=name, email=email, role=Role.TEACHER, faculty_id=faculty.id, program_id=program.id)
            teachers.append(get_or_create_teacher(session, user, staff_no))

        plos = ensure_me_plos(session, program)
        courses = ensure_courses(session, program)
        course_by_ysem: dict[tuple[int, str], list[Course]] = {}
        for course in courses:
            course_by_ysem.setdefault((course.curriculum_year or 1, str(course.curriculum_semester or "1")), []).append(course)

        for year in range(1, 5):
            source_students = ensure_students_for_year(session, program, faculty, year)
            academic_year = YEAR_TO_ACADEMIC_YEAR[year]
            for semester in (1, 2):
                code = class_code(year, semester)
                semester_label = f"Semester {semester}"
                for course in course_by_ysem.get((year, str(semester)), []):
                    course_class = session.exec(select(CourseClass).where(CourseClass.course_id == course.id, CourseClass.name == code, CourseClass.academic_year == academic_year)).first()
                    if not course_class:
                        course_class = CourseClass(course_id=course.id, academic_year=academic_year, semester=semester_label, name=code, semester_start=f"{academic_year[:4]}-08-01", semester_end=f"{academic_year[-4:]}-07-31")
                        session.add(course_class)
                    else:
                        course_class.semester = semester_label
                        course_class.semester_start = f"{academic_year[:4]}-08-01"
                        course_class.semester_end = f"{academic_year[-4:]}-07-31"
                        session.add(course_class)
                    session.commit()
                    session.refresh(course_class)

                    teacher = teachers[((year - 1) * 2 + (semester - 1)) % len(teachers)]
                    ensure_assignment(session, course, course_class, teacher)
                    assessments = ensure_assessment_plan(session, course, plos)

                    # Enrol demo students only in Semester 1 of each year. Semester 2 exists as promotion target.
                    if semester == 1:
                        for student in source_students:
                            ensure_class_student(session, course_class, student)
                            ensure_semester_enrollment(session, program, student, code, academic_year, semester_label)
                        seed_scores_for_teacher(session, source_students, assessments, teacher.user_id, seed_value=(year * 1000 + course.id))

        session.add(AuditLog(
            date_time=datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
            user_name="Seed Script",
            module="Demo Data",
            action="Create/Update",
            description="Seeded ME teachers, role assignments, cohorts, students, assessments, and teacher-entered scores.",
            item_record="ME teacher score demo",
            ip_address="127.0.0.1",
            status="Success",
        ))
        session.commit()
        print("ME teacher demo data created successfully.")
        print("Login examples:")
        print("  Programme Coordinator: me.manager@example.com / password")
        for _, name, email in TEACHERS:
            print(f"  Teacher: {email} / password")
        print("Student example: student.21me101@example.com / password")


if __name__ == "__main__":
    run()

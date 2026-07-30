from datetime import datetime
from enum import Enum
from typing import Optional

from sqlalchemy import UniqueConstraint
from sqlmodel import Field, Relationship, SQLModel


class Role(str, Enum):
    SUPER_ADMIN = "super_admin"
    DEAN = "dean"
    PROGRAM_MANAGER = "program_manager"
    TEACHER = "teacher"
    STUDENT = "student"


class RoleDefinition(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("role_code", name="uq_role_definition_code"),)

    id: int | None = Field(default=None, primary_key=True)
    role_key: str | None = Field(default=None, index=True)
    role_name: str
    role_code: str = Field(index=True)
    description: str = ""
    status: str = "Active"
    is_system_role: bool = False
    abac_scope_type: str = "All"
    menu_access: str = ""
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None


class RolePermission(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("role_definition_id", "module", name="uq_role_permission_module"),)

    id: int | None = Field(default=None, primary_key=True)
    role_definition_id: int = Field(foreign_key="roledefinition.id")
    module: str
    can_view: bool = False
    can_create: bool = False
    can_edit: bool = False
    can_delete: bool = False
    can_export: bool = False
    updated_at: datetime | None = None


class User(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str
    email: str = Field(index=True, unique=True)
    password_hash: str
    role: Role
    is_active: bool = True
    avatar_name: str = ""
    faculty_id: int | None = Field(default=None, foreign_key="faculty.id")
    program_id: int | None = Field(default=None, foreign_key="program.id")

    teacher_profile: Optional["Teacher"] = Relationship(back_populates="user")
    student_profile: Optional["Student"] = Relationship(back_populates="user")


class Faculty(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str

    programs: list["Program"] = Relationship(back_populates="faculty")


class Program(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    faculty_id: int = Field(foreign_key="faculty.id")
    name: str
    code: str

    faculty: Faculty = Relationship(back_populates="programs")
    plos: list["PLO"] = Relationship(back_populates="program")
    courses: list["Course"] = Relationship(back_populates="program")


class AcademicYear(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str = Field(index=True, unique=True)
    start_date: str
    end_date: str
    is_active: bool = True
    is_default: bool = False
    created_by: str = "Admin"
    created_at: str = "May 15, 2025 10:15 AM"


class AcademicSemester(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str
    code: str = Field(index=True, unique=True)
    academic_year: str
    start_date: str
    end_date: str
    is_active: bool = True
    is_default: bool = False
    created_at: str = "Apr 10, 2025 09:15 AM"


class Teacher(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", unique=True)
    staff_no: str

    user: User = Relationship(back_populates="teacher_profile")
    class_assignments: list["ClassTeacher"] = Relationship(back_populates="teacher")
    course_assignments: list["CourseTeacher"] = Relationship(back_populates="teacher")


class Student(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("user_id", name="uq_student_user_account"),)

    id: int | None = Field(default=None, primary_key=True)
    user_id: int | None = Field(default=None, foreign_key="user.id")
    student_no: str = Field(index=True, unique=True)
    name_kh: str | None = None
    name_en: str

    user: Optional[User] = Relationship(back_populates="student_profile")
    enrollments: list["ClassStudent"] = Relationship(back_populates="student")
    scores: list["StudentScore"] = Relationship(back_populates="student")


class Course(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    program_id: int = Field(foreign_key="program.id")
    cohort_id: int | None = Field(default=None, foreign_key="courseclass.id")
    code: str
    title: str
    credits: float = 3
    curriculum_year: int | None = None
    curriculum_semester: str | None = None

    program: Program = Relationship(back_populates="courses")
    clos: list["CLO"] = Relationship(back_populates="course")
    teacher_assignments: list["CourseTeacher"] = Relationship(back_populates="course")


class CourseTeacher(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    course_id: int = Field(foreign_key="course.id")
    teacher_id: int = Field(foreign_key="teacher.id")

    course: Course = Relationship(back_populates="teacher_assignments")
    teacher: Teacher = Relationship(back_populates="course_assignments")


class CourseClass(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    course_id: int = Field(foreign_key="course.id")
    academic_year: str
    semester: str
    semester_start: str | None = None
    semester_end: str | None = None
    name: str

    course: Course = Relationship(sa_relationship_kwargs={"foreign_keys": "CourseClass.course_id"})
    teachers: list["ClassTeacher"] = Relationship(back_populates="course_class")
    students: list["ClassStudent"] = Relationship(back_populates="course_class", sa_relationship_kwargs={"foreign_keys": "ClassStudent.class_id"})


class ClassTeacher(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    class_id: int = Field(foreign_key="courseclass.id")
    teacher_id: int = Field(foreign_key="teacher.id")

    course_class: CourseClass = Relationship(back_populates="teachers")
    teacher: Teacher = Relationship(back_populates="class_assignments")


class ClassStudent(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    class_id: int = Field(foreign_key="courseclass.id")
    student_id: int = Field(foreign_key="student.id")
    status: str = "Active"
    promoted_to_class_id: int | None = Field(default=None, foreign_key="courseclass.id")
    promoted_at: datetime | None = None
    promoted_by_user_id: int | None = Field(default=None, foreign_key="user.id")

    course_class: CourseClass = Relationship(back_populates="students", sa_relationship_kwargs={"foreign_keys": "ClassStudent.class_id"})
    student: Student = Relationship(back_populates="enrollments")


class StudentSemesterEnrollment(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("student_id", "program_id", "cohort_name", "academic_year", "semester", name="uq_student_semester_enrollment"),)

    id: int | None = Field(default=None, primary_key=True)
    student_id: int = Field(foreign_key="student.id")
    program_id: int = Field(foreign_key="program.id")
    cohort_name: str
    academic_year: str
    semester: str
    status: str = "Active"
    promoted_from_id: int | None = Field(default=None, foreign_key="studentsemesterenrollment.id")
    promoted_by_user_id: int | None = Field(default=None, foreign_key="user.id")
    promoted_at: datetime | None = None
    created_at: datetime = Field(default_factory=datetime.utcnow)


class StudentPromotionHistory(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    promoted_by_user_id: int | None = Field(default=None, foreign_key="user.id")
    promoted_by_name: str = "Admin"
    academic_year: str
    faculty_id: int | None = Field(default=None, foreign_key="faculty.id")
    program_id: int | None = Field(default=None, foreign_key="program.id")
    cohort_name: str
    from_semester: str
    to_semester: str
    student_count: int = 0
    course_count: int = 0
    skipped_count: int = 0
    created_at: datetime = Field(default_factory=datetime.utcnow)


class PLOVersion(SQLModel, table=True):
    __tablename__ = "plo_versions"
    __table_args__ = (UniqueConstraint("programme_id", "version_name", name="uq_plo_version_programme_name"),)

    id: int | None = Field(default=None, primary_key=True)
    programme_id: int = Field(foreign_key="program.id")
    version_name: str
    effective_academic_year_id: int | None = Field(default=None, foreign_key="academicyear.id")
    status: str = "Active"
    is_locked: bool = False
    created_by: int | None = Field(default=None, foreign_key="user.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None




class CohortOutcomeVersion(SQLModel, table=True):
    """Assign one shared PEO/PLO/programme-mapping version to a cohort."""
    __tablename__ = "cohort_outcome_versions"
    __table_args__ = (UniqueConstraint("programme_id", "cohort_name", name="uq_cohort_outcome_version"),)

    id: int | None = Field(default=None, primary_key=True)
    programme_id: int = Field(foreign_key="program.id")
    cohort_name: str = Field(index=True)
    outcome_version_id: int = Field(foreign_key="plo_versions.id")
    assigned_by: int | None = Field(default=None, foreign_key="user.id")
    assigned_at: datetime = Field(default_factory=datetime.utcnow)

class PEO(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    program_id: int = Field(foreign_key="program.id")
    plo_version_id: int | None = Field(default=None, foreign_key="plo_versions.id")
    code: str
    description: str
    status: str = "Active"
    remark: str = ""
    created_by: int | None = Field(default=None, foreign_key="user.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None


class PLO(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    program_id: int = Field(foreign_key="program.id")
    plo_version_id: int | None = Field(default=None, foreign_key="plo_versions.id")
    code: str
    description: str
    domain: str = "Knowledge"
    bloom_level: str = "C1"
    status: str = "Active"
    remark: str = ""
    created_by: int | None = Field(default=None, foreign_key="user.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None

    program: Program = Relationship(back_populates="plos")
    mappings: list["CLOPLOMapping"] = Relationship(back_populates="plo")


class CLO(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    course_id: int = Field(foreign_key="course.id")
    code: str
    domain: str = "K,S,A"
    description: str
    pass_threshold: float = 0.5

    course: Course = Relationship(back_populates="clos")
    mappings: list["CLOPLOMapping"] = Relationship(back_populates="clo")
    assessments: list["Assessment"] = Relationship(back_populates="clo")


class CLOPLOMapping(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    clo_id: int = Field(foreign_key="clo.id")
    plo_id: int = Field(foreign_key="plo.id")
    weight: float = 0

    clo: CLO = Relationship(back_populates="mappings")
    plo: PLO = Relationship(back_populates="mappings")


class CoursePLOMapping(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    course_id: int = Field(foreign_key="course.id")
    plo_id: int = Field(foreign_key="plo.id")
    level: int = 0
    symbol: str = ""


class PEOPLOMapping(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    program_id: int | None = Field(default=None, foreign_key="program.id")
    plo_version_id: int | None = Field(default=None, foreign_key="plo_versions.id")
    peo_id: int = Field(foreign_key="peo.id")
    plo_id: int = Field(foreign_key="plo.id")
    mapping_mode: str = "percentage"
    is_mapped: bool = True
    contribution_percentage: float = 0
    created_by: int | None = Field(default=None, foreign_key="user.id")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None
    weight: float = 0


class Assessment(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    clo_id: int = Field(foreign_key="clo.id")
    name: str
    description: str | None = None
    max_score: float
    weight: float = 1

    clo: CLO = Relationship(back_populates="assessments")
    scores: list["StudentScore"] = Relationship(back_populates="assessment")


class StudentScore(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("assessment_id", "student_id", name="uq_student_assessment_score"),)

    id: int | None = Field(default=None, primary_key=True)
    assessment_id: int = Field(foreign_key="assessment.id")
    student_id: int = Field(foreign_key="student.id")
    score: float = 0
    status: str = "Draft"
    locked: bool = False
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None
    submitted_at: datetime | None = None
    submitted_by_user_id: int | None = Field(default=None, foreign_key="user.id")
    entered_by_user_id: int | None = Field(default=None, foreign_key="user.id")

    assessment: Assessment = Relationship(back_populates="scores")
    student: Student = Relationship(back_populates="scores")


class PLOTarget(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    program_id: int = Field(foreign_key="program.id")
    plo_id: int = Field(foreign_key="plo.id")
    academic_year: str = "2024-2025"
    cohort: str = "Cohort 2024"
    target: float = 70
    set_by: str = "Admin"
    updated_at: str = "Apr 15, 2024 10:15 AM"


class SystemReport(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str
    category: str
    description: str
    created_by: str = "Admin"
    last_generated: str = "May 15, 2024 10:30 AM"
    format: str = "PDF"
    status: str = "Ready"


class AuditLog(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    date_time: str
    user_name: str = "Admin"
    module: str
    action: str
    description: str
    item_record: str
    ip_address: str = "192.168.1.10"
    status: str = "Success"


class SystemSetting(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    key: str = Field(index=True, unique=True)
    value: str
    category: str = "general"


class StudyPeriod(SQLModel, table=True):
    """Global academic context shared by all roles."""
    __table_args__ = (
        UniqueConstraint("academic_year", "semester", name="uq_study_period_year_semester"),
    )

    id: int | None = Field(default=None, primary_key=True)
    academic_year: str = Field(index=True)
    semester: int = Field(index=True)
    label: str
    start_date: str | None = None
    end_date: str | None = None
    is_active: bool = True
    is_current: bool = False
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None


class UserStudyPeriodPreference(SQLModel, table=True):
    """Persists the last study period selected by each user."""
    __table_args__ = (
        UniqueConstraint("user_id", name="uq_user_study_period_preference_user"),
    )

    id: int | None = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    study_period_id: int = Field(foreign_key="studyperiod.id", index=True)
    updated_at: datetime = Field(default_factory=datetime.utcnow)


class Document(SQLModel, table=True):
    """An uploaded document, scoped to a faculty and/or programme.

    Both scope columns empty means the document is visible system-wide.
    """

    id: int | None = Field(default=None, primary_key=True)
    title: str
    description: str = ""
    category: str = "General"
    status: str = "Published"

    original_name: str = ""
    stored_name: str = ""
    content_type: str = ""
    size_bytes: int = 0

    faculty_id: int | None = Field(default=None, foreign_key="faculty.id", index=True)
    program_id: int | None = Field(default=None, foreign_key="program.id", index=True)

    uploaded_by: int | None = Field(default=None, foreign_key="user.id")
    uploaded_by_name: str = ""
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime | None = None


class UserProgrammePreference(SQLModel, table=True):
    """Persists the programme a Dean is currently working on.

    Deans manage every programme in their faculty, so the manager pages need to
    know which one is active. This is kept separate from `User.program_id`, which
    stays empty for Deans and means "belongs to this programme".
    """
    __table_args__ = (
        UniqueConstraint("user_id", name="uq_user_programme_preference_user"),
    )

    id: int | None = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    program_id: int = Field(foreign_key="program.id", index=True)
    updated_at: datetime = Field(default_factory=datetime.utcnow)

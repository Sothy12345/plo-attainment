from typing import Iterable

from app.models import Program, Role, User


ROLE_LABELS = {
    Role.SUPER_ADMIN: "Admin",
    Role.DEAN: "Dean",
    Role.PROGRAM_MANAGER: "Program Manager",
    Role.TEACHER: "Teacher",
    Role.STUDENT: "Student",
}

PERMISSIONS = {
    "manage_system": {Role.SUPER_ADMIN},
    "view_all_reports": {Role.SUPER_ADMIN, Role.DEAN},
    # NOTE: manage_program / manage_clo_assessment gate the legacy /setup pages,
    # which list every programme in the system with no faculty filter. Deans are
    # deliberately left out until those pages are scoped.
    "manage_program": {Role.SUPER_ADMIN, Role.PROGRAM_MANAGER},
    "promote_students": {Role.SUPER_ADMIN, Role.PROGRAM_MANAGER, Role.DEAN},
    "manage_clo_assessment": {Role.SUPER_ADMIN, Role.PROGRAM_MANAGER},
    "input_marks": {Role.SUPER_ADMIN, Role.TEACHER},
    "import_marks": {Role.SUPER_ADMIN},
    "view_clo_report": {Role.SUPER_ADMIN, Role.DEAN, Role.PROGRAM_MANAGER, Role.TEACHER},
    "view_own_report": {Role.STUDENT},
}


def can(user: User, permission: str) -> bool:
    return user.role in PERMISSIONS.get(permission, set())


def scope_label(user: User) -> str:
    if user.role == Role.SUPER_ADMIN:
        return "All faculties and all programmes"
    if user.role == Role.DEAN:
        return "Assigned faculty and all programmes under that faculty"
    if user.role == Role.PROGRAM_MANAGER:
        return "Assigned programme"
    if user.role == Role.TEACHER:
        return "Assigned courses/classes, including cross-programme teaching assignments"
    return "Own student record and enrolled classes"


def can_access_faculty(user: User, faculty_id: int | None) -> bool:
    if user.role == Role.SUPER_ADMIN:
        return True
    return faculty_id is not None and user.faculty_id == faculty_id


def can_access_program(user: User, program: Program | None) -> bool:
    if not program:
        return False
    if user.role == Role.SUPER_ADMIN:
        return True
    if user.role == Role.DEAN:
        return user.faculty_id is not None and program.faculty_id == user.faculty_id
    if user.role in {Role.PROGRAM_MANAGER, Role.TEACHER, Role.STUDENT}:
        return user.program_id is not None and program.id == user.program_id
    return False


def scoped_programs(user: User, programs: Iterable[Program]) -> list[Program]:
    return [program for program in programs if can_access_program(user, program)]

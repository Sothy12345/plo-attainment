from sqlmodel import Session, select

from app.models import (
    Assessment,
    CLO,
    CLOPLOMapping,
    CoursePLOMapping,
    ClassStudent,
    ClassTeacher,
    Course,
    CourseClass,
    Faculty,
    PEO,
    PEOPLOMapping,
    PLO,
    Program,
    Role,
    Student,
    StudentScore,
    Teacher,
    User,
)
from app.security import hash_password


ME_PEOS = [
    ("PEO1", "To produce qualified engineers in Mechanical engineering with industrial engineering to meet market needs in the south-east region of Cambodia."),
    ("PEO2", "To produce life-long learners who contribute significantly to the mechanical and industrial engineering field."),
    ("PEO3", "To pursue post graduate education in the field mechanical and industrial engineering."),
    ("PEO4", "To produce self-intrapreneurs in the fields related to mechanical and industrial engineering with professional skills."),
]

ME_PLOS = [
    ("PLO1", "Apply engineering knowledge to solve complex problems in the mechanical and industrial fields."),
    ("PLO2", "Design the mechanical and industrial engineering systems to meet specific needs."),
    ("PLO3", "Conduct the practical skills by using the mechanical tools and machinery competently."),
    ("PLO4", "Develop effective teamwork whose members together provide collaboration from diverse disciplines to achieve common goals."),
    ("PLO5", "Demonstrate professional responsibility in mechanical and industrial engineering work which should consider the global impact, environmentally friendly, and societal contexts."),
    ("PLO6", "Analyse business problems across a wide range of business domains."),
    ("PLO7", "Develop ethics and professional responsibility in mechanical and industrial context which consider societal impact."),
    ("PLO8", "Apply holistic communication with scientific and non-scientific communities by using appropriate technical literature."),
    ("PLO9", "Use computer software for data analysis and simulation to solve complex problems in the context of mechanical and industrial engineering."),
    ("PLO10", "Use numerical techniques and tools to solve engineering problems."),
]

FACULTY_PROGRAM_STRUCTURE = [
    (
        "FBA",
        "Faculty of Business Administration",
        [
            ("ACC", "Accounting"),
            ("MGT", "Management"),
            ("MKT", "Marketing"),
            ("FBK", "Finance and Banking"),
        ],
    ),
    (
        "FST",
        "Faculty of Science and Technology",
        [
            ("CS", "Computer Science"),
            ("MATH", "Mathematics"),
            ("ME", "Mechanical Engineering"),
            ("EAE", "Electronics and Automation Engineering"),
        ],
    ),
    (
        "FA",
        "Faculty of Agriculture",
        [
            ("AGRO", "Agronomy"),
            ("ASV", "Animal Science and Veterinary"),
            ("RD", "Rural Development"),
        ],
    ),
    (
        "FAHFL",
        "Faculty of Art, Humanity and Foreign Language",
        [
            ("TEFL", "English Language / TEFL"),
        ],
    ),
    (
        "FSS",
        "Faculty of Social Sciences",
        [
            ("PA", "Public Administration"),
        ],
    ),
]


def sync_faculty_program_structure(session: Session) -> None:
    faculties_by_name = {faculty.name: faculty for faculty in session.exec(select(Faculty)).all()}
    programs_by_code = {program.code: program for program in session.exec(select(Program)).all()}

    for faculty_code, faculty_name, programmes in FACULTY_PROGRAM_STRUCTURE:
        faculty = faculties_by_name.get(faculty_name)
        if not faculty:
            faculty = Faculty(name=faculty_name)
            session.add(faculty)
            session.commit()
            session.refresh(faculty)
            faculties_by_name[faculty_name] = faculty

        for program_code, program_name in programmes:
            program = programs_by_code.get(program_code)
            if program:
                program.name = program_name
                program.faculty_id = faculty.id
                session.add(program)
            else:
                program = Program(code=program_code, name=program_name, faculty_id=faculty.id)
                session.add(program)
                session.flush()
                programs_by_code[program_code] = program
    session.commit()

    legacy_faculty = session.exec(select(Faculty).where(Faculty.name == "Faculty of Engineering")).first()
    if legacy_faculty:
        linked_program = session.exec(select(Program).where(Program.faculty_id == legacy_faculty.id)).first()
        if not linked_program:
            session.delete(legacy_faculty)
            session.commit()

    fst = session.exec(select(Faculty).where(Faculty.name == "Faculty of Science and Technology")).first()
    me = session.exec(select(Program).where(Program.code == "ME")).first()
    if fst and me:
        role_scopes = {
            "dean@example.com": (fst.id, None),
            "manager@example.com": (fst.id, me.id),
            "teacher@example.com": (fst.id, me.id),
            "teacher2@example.com": (fst.id, me.id),
            "student@example.com": (fst.id, me.id),
        }
        for email, (faculty_id, program_id) in role_scopes.items():
            user = session.exec(select(User).where(User.email == email)).first()
            if user:
                user.faculty_id = faculty_id
                user.program_id = program_id
                session.add(user)
        session.commit()

ME_CURRICULUM = [
    (1, "1", "2SE-ME001", "General English 1A", 3),
    (1, "1", "2SE-ME002", "Mathematics for Engineering 1", 3),
    (1, "1", "2SE-ME003", "Geometry for Engineering", 2),
    (1, "1", "2SE-ME004", "Engineering Mechanics 1", 2.5),
    (1, "1", "2SE-ME005", "CSDGs and ASEAN", 3),
    (1, "1", "ME-Y1S1-REL", "Communication and Interpersonal Relation", 1.5),
    (1, "1", "2SE-ME006", "Digital literacy*", 3),
    (1, "2", "2SE-ME007", "General English 1B", 3),
    (1, "2", "2SE-ME008", "Mathematics for Engineering 2", 3),
    (1, "2", "2SE-ME009", "Engineering Mechanics 2", 2.5),
    (1, "2", "2SE-ME010", "Thermodynamics", 2),
    (1, "2", "2SE-ME011", "Computer Aided Design (Autocad/ Solidwork)", 3),
    (1, "2", "ME-Y1S2-UCL", "University and Community Linkage", 1.5),
    (1, "2", "2SE-ME012", "Technical Drawing*", 2),
    (2, "1", "2SE-ME013", "Welding Techniques", 3),
    (2, "1", "2SE-ME014", "Applied Thermodynamics", 2),
    (2, "1", "2SE-ME015", "General English 2A", 3),
    (2, "1", "2SE-ME016", "Solid Mechanics", 3),
    (2, "1", "2SE-ME017", "Probability and Statistic", 3),
    (2, "2", "2SE-ME018", "Materials Science", 3),
    (2, "2", "ME-Y2S2-ELEC", "Electricity", 2),
    (2, "2", "2SE-ME020", "Strength of Materials", 3),
    (2, "2", "2SE-ME021", "Fluid Mechanics", 3),
    (2, "2", "2SE-ME022", "General English 2B", 3),
    (2, "2", "2SE-ME023", "Heat Transfer", 2),
    (2, "2", "2SE-ME024", "Mechanical Design", 3),
    (2, "2", "2SE-ME025", "Numerical Method and Optimization", 2),
    (2, "2", "ME-Y2S2-EM", "Electric Machine", 2),
    (2, "2", "ME-Y2S2-FDWM", "Facility Design and Warehouse Management", 2),
    (2, "2", "ME-Y2S2-IRM", "Introduction to Research Methodology", 2),
    (3, "1", "2SE-ME027", "Mechanical Vibration", 2),
    (3, "1", "2SE-ME028", "Heat Exchangers", 2),
    (3, "1", "2SE-ME030", "Internal Combustion Engine 1", 2),
    (3, "1", "2SE-ME031", "Industrial Hydraulics and Pneumatics Systems", 2),
    (3, "1", "ME-Y3S1-LM", "Logistic Management", 2),
    (3, "1", "ME-Y3S1-LCT", "Logistic and Critical Thinking", 2),
    (3, "1", "2SE-ME032", "Refrigeration and Air Conditioning System 1", 3),
    (3, "2", "2SE-ME033", "Machine elements", 3),
    (3, "2", "2SE-ME034", "Mechanical Production", 3),
    (3, "2", "2SE-ME035", "Internal Combustion Engine 2", 2),
    (3, "2", "2SE-ME036", "Refrigeration and Air Conditioning System 2", 3),
    (3, "2", "2SE-ME037", "Turbomachines", 2),
    (3, "2", "ME-Y3S2-P1", "Mechanical/Industrial Project I", 1),
    (3, "2", "2SE-ME039", "Finite Element Analysis", 2),
    (3, "2", "2SE-ME040", "Internship (4-6 weeks)", 2),
    (4, "1", "2SE-ME041", "Data Analysis", 2),
    (4, "1", "2SE-ME042", "Project Management", 2),
    (4, "1", "2SE-ME043", "Lean Manufacturing/Production Engineering", 3),
    (4, "1", "2SE-ME044", "Entrepreneurship", 2),
    (4, "1", "ME-Y4S1-P2", "Mechanical/Industrial Project II", 1),
    (4, "1", "2SE-ME045", "Boiler Technology", 2),
    (4, "2", "2SE-ME046", "Final Year Internship", 9),
]


ME_PEO_PLO_MAPPING = {
    "PEO1": ["PLO1", "PLO2", "PLO3", "PLO4", "PLO5", "PLO6", "PLO7", "PLO8", "PLO9", "PLO10"],
    "PEO2": ["PLO1", "PLO2", "PLO3", "PLO4", "PLO5", "PLO6", "PLO7", "PLO8", "PLO9", "PLO10"],
    "PEO3": ["PLO1", "PLO2", "PLO3", "PLO4", "PLO5", "PLO6", "PLO7", "PLO8", "PLO9", "PLO10"],
    "PEO4": ["PLO1", "PLO4", "PLO5", "PLO6", "PLO7", "PLO8", "PLO9", "PLO10"],
}


ME_COURSE_PLO_MAPPING = [
    ("2SE-ME001", ["P", "", "", "M", "", "", "", "M", "", ""]),
    ("2SE-ME002", ["P", "M", "", "", "", "", "", "", "", "M"]),
    ("2SE-ME003", ["P", "P", "", "", "", "", "", "", "", "M"]),
    ("2SE-ME004", ["P", "", "P", "", "", "", "", "", "", "M"]),
    ("2SE-ME005", ["P", "P", "", "P", "", "", "P", "P", "", ""]),
    ("ME-Y1S1-REL", ["P", "", "", "P", "P", "", "", "M", "", ""]),
    ("2SE-ME006", ["P", "P", "", "", "", "", "", "", "F", ""]),
    ("2SE-ME007", ["P", "", "", "M", "", "", "", "M", "", ""]),
    ("2SE-ME008", ["P", "M", "", "", "", "", "", "", "", "M"]),
    ("2SE-ME009", ["P", "", "P", "", "", "", "", "", "", "M"]),
    ("2SE-ME010", ["P", "", "P", "", "", "", "", "", "", "M"]),
    ("2SE-ME011", ["M", "", "P", "", "P", "", "", "", "P", ""]),
    ("ME-Y1S2-UCL", ["P", "", "", "", "P", "", "P", "P", "", ""]),
    ("2SE-ME012", ["P", "P", "M", "", "", "", "", "", "", ""]),
    ("2SE-ME013", ["P", "", "F", "", "", "P", "", "", "", ""]),
    ("2SE-ME014", ["M", "M", "M", "P", "", "", "", "P", "", ""]),
    ("2SE-ME015", ["P", "", "", "M", "", "", "", "M", "", ""]),
    ("2SE-ME016", ["P", "", "P", "", "P", "", "", "", "P", ""]),
    ("2SE-ME017", ["P", "M", "", "", "", "", "", "", "", "M"]),
    ("2SE-ME018", ["P", "M", "P", "", "", "", "", "", "", ""]),
    ("ME-Y2S2-ELEC", ["P", "", "P", "", "P", "", "P", "", "", ""]),
    ("2SE-ME020", ["F", "", "P", "", "", "", "P", "", "", ""]),
    ("2SE-ME021", ["F", "", "P", "", "", "", "", "", "", "P"]),
    ("2SE-ME022", ["P", "", "", "M", "", "", "", "M", "", ""]),
    ("2SE-ME023", ["P", "M", "P", "", "", "", "", "P", "", ""]),
    ("2SE-ME024", ["P", "", "M", "", "", "", "P", "", "P", ""]),
    ("2SE-ME025", ["P", "P", "", "", "", "", "", "", "P", "P"]),
    ("ME-Y2S2-EM", ["F", "P", "P", "", "", "", "", "", "", ""]),
    ("ME-Y2S2-FDWM", ["P", "M", "", "", "P", "", "", "", "", ""]),
    ("ME-Y2S2-IRM", ["P", "", "", "M", "M", "", "P", "", "", ""]),
    ("2SE-ME027", ["P", "", "P", "", "", "", "", "", "", "M"]),
    ("2SE-ME028", ["P", "M", "P", "", "", "", "", "P", "", ""]),
    ("2SE-ME030", ["F", "", "M", "", "", "P", "", "", "", ""]),
    ("2SE-ME031", ["M", "P", "P", "", "", "", "", "", "", ""]),
    ("ME-Y3S1-LM", ["F", "P", "", "P", "", "P", "", "", "", ""]),
    ("2SE-ME025", ["F", "P", "P", "", "", "", "", "", "", "P"]),
    ("2SE-ME032", ["P", "", "P", "", "", "P", "", "", "P", "P"]),
    ("2SE-ME033", ["P", "P", "M", "", "", "", "", "", "", ""]),
    ("2SE-ME034", ["M", "M", "", "", "", "", "", "", "", "P"]),
    ("2SE-ME035", ["M", "M", "", "", "", "", "", "", "", "P"]),
    ("2SE-ME036", ["P", "", "P", "", "", "P", "", "", "P", "P"]),
    ("2SE-ME037", ["F", "P", "P", "", "", "", "", "", "", ""]),
    ("ME-Y3S2-P1", ["P", "P", "P", "P", "P", "P", "P", "P", "P", "P"]),
    ("2SE-ME039", ["P", "P", "M", "", "", "", "", "", "", ""]),
    ("2SE-ME040", ["P", "P", "P", "P", "P", "P", "P", "P", "P", "P"]),
    ("2SE-ME041", ["P", "P", "", "", "", "", "M", "", "", ""]),
    ("2SE-ME042", ["F", "P", "", "", "", "", "", "", "P", "P"]),
    ("2SE-ME043", ["M", "M", "", "", "", "", "", "", "", "P"]),
    ("2SE-ME044", ["P", "", "", "", "P", "F", "", "", "", ""]),
    ("ME-Y4S1-P2", ["P", "P", "P", "P", "P", "P", "P", "P", "P", "P"]),
    ("2SE-ME045", ["P", "M", "", "", "", "P", "", "", "", ""]),
    ("2SE-ME046", ["P", "P", "P", "P", "P", "P", "P", "P", "P", "P"]),
]


def mapping_level(symbol: str) -> int:
    return {"M": 100, "F": 60, "P": 30}.get(symbol.strip().upper(), 0)


def mapping_symbol(symbol: str) -> str:
    level = mapping_level(symbol)
    return f"{level}%" if level else ""


def sync_me_specification(session: Session) -> None:
    sync_faculty_program_structure(session)
    program = session.exec(select(Program).where(Program.code == "ME")).first()
    if not program:
        return

    second_teacher_user = session.exec(select(User).where(User.email == "teacher2@example.com")).first()
    if not second_teacher_user:
        second_teacher_user = User(name="Sok Dara", email="teacher2@example.com", password_hash=hash_password("password"), role=Role.TEACHER)
        session.add(second_teacher_user)
        session.commit()
        session.refresh(second_teacher_user)
    fst = session.exec(select(Faculty).where(Faculty.name == "Faculty of Science and Technology")).first()
    if fst and program:
        second_teacher_user.faculty_id = fst.id
        second_teacher_user.program_id = program.id
        session.add(second_teacher_user)
    second_teacher = session.exec(select(Teacher).where(Teacher.user_id == second_teacher_user.id)).first()
    if not second_teacher:
        session.add(Teacher(user_id=second_teacher_user.id, staff_no="T-002"))

    for code, description in ME_PEOS:
        peo = session.exec(select(PEO).where(PEO.program_id == program.id, PEO.code == code)).first()
        if peo:
            peo.description = description
            session.add(peo)
        else:
            session.add(PEO(program_id=program.id, code=code, description=description))

    for code, description in ME_PLOS:
        plo = session.exec(select(PLO).where(PLO.program_id == program.id, PLO.code == code)).first()
        if plo:
            plo.description = description
            session.add(plo)
        else:
            session.add(PLO(program_id=program.id, code=code, description=description))

    official_codes = {code for code, _ in ME_PLOS}
    extra_plos = session.exec(select(PLO).where(PLO.program_id == program.id)).all()
    for plo in extra_plos:
        has_mapping = session.exec(select(CLOPLOMapping).where(CLOPLOMapping.plo_id == plo.id)).first()
        if plo.code not in official_codes and not has_mapping:
            session.delete(plo)

    legacy_cad = session.exec(select(Course).where(Course.code == "2SE-ME012", Course.title == "Computer Aided Design")).first()
    if legacy_cad:
        legacy_cad.code = "2SE-ME011"
        legacy_cad.title = "Computer Aided Design (Autocad/ Solidwork)"
        legacy_cad.credits = 3
        legacy_cad.curriculum_year = 1
        legacy_cad.curriculum_semester = "2"
        session.add(legacy_cad)

    for year, semester, code, title, credits in ME_CURRICULUM:
        course = session.exec(select(Course).where(Course.program_id == program.id, Course.code == code)).first()
        if course:
            course.title = title
            course.credits = credits
            course.curriculum_year = year
            course.curriculum_semester = semester
            session.add(course)
        else:
            session.add(
                Course(
                    program_id=program.id,
                    code=code,
                    title=title,
                    credits=credits,
                    curriculum_year=year,
                    curriculum_semester=semester,
                )
            )
    session.commit()

    for _year, _semester, code, _title, _credits in ME_CURRICULUM:
        duplicates = sorted(
            session.exec(select(Course).where(Course.program_id == program.id, Course.code == code)).all(),
            key=lambda item: item.id or 0,
        )
        if len(duplicates) <= 1:
            continue
        keeper = duplicates[0]
        for candidate in duplicates:
            has_class = session.exec(select(CourseClass).where(CourseClass.course_id == candidate.id)).first()
            has_clo = session.exec(select(CLO).where(CLO.course_id == candidate.id)).first()
            if has_class or has_clo:
                keeper = candidate
                break
        for duplicate in duplicates:
            if duplicate.id == keeper.id:
                continue
            for mapping in session.exec(select(CoursePLOMapping).where(CoursePLOMapping.course_id == duplicate.id)).all():
                existing = session.exec(
                    select(CoursePLOMapping).where(
                        CoursePLOMapping.course_id == keeper.id,
                        CoursePLOMapping.plo_id == mapping.plo_id,
                    )
                ).first()
                if existing:
                    session.delete(mapping)
                else:
                    mapping.course_id = keeper.id
                    session.add(mapping)
            has_class = session.exec(select(CourseClass).where(CourseClass.course_id == duplicate.id)).first()
            has_clo = session.exec(select(CLO).where(CLO.course_id == duplicate.id)).first()
            if not has_class and not has_clo:
                session.delete(duplicate)
        session.commit()

    peos_by_code = {peo.code: peo for peo in session.exec(select(PEO).where(PEO.program_id == program.id)).all()}
    plos_by_code = {plo.code: plo for plo in session.exec(select(PLO).where(PLO.program_id == program.id)).all()}
    courses_by_code = {course.code: course for course in session.exec(select(Course).where(Course.program_id == program.id)).all()}

    for peo_code, plo_codes in ME_PEO_PLO_MAPPING.items():
        peo = peos_by_code.get(peo_code)
        if not peo:
            continue
        existing_links = session.exec(select(PEOPLOMapping).where(PEOPLOMapping.peo_id == peo.id)).all()
        existing_plo_ids = {link.plo_id for link in existing_links}
        for plo_code in plo_codes:
            plo = plos_by_code.get(plo_code)
            if plo and plo.id not in existing_plo_ids:
                session.add(PEOPLOMapping(peo_id=peo.id, plo_id=plo.id))

    for course_code, symbols in ME_COURSE_PLO_MAPPING:
        course = courses_by_code.get(course_code)
        if not course:
            continue
        for index, symbol in enumerate(symbols, 1):
            plo = plos_by_code.get(f"PLO{index}")
            if not plo:
                continue
            mapping = session.exec(
                select(CoursePLOMapping).where(CoursePLOMapping.course_id == course.id, CoursePLOMapping.plo_id == plo.id)
            ).first()
            level = mapping_level(symbol)
            display_symbol = mapping_symbol(symbol)
            if mapping:
                mapping.level = level
                mapping.symbol = display_symbol
                session.add(mapping)
            else:
                session.add(CoursePLOMapping(course_id=course.id, plo_id=plo.id, level=level, symbol=display_symbol))
    for mapping in session.exec(select(CLOPLOMapping)).all():
        if 0 < mapping.weight <= 1:
            mapping.weight = round(mapping.weight * 100, 2)
            session.add(mapping)
    session.commit()


def seed_data(session: Session) -> None:
    existing = session.exec(select(User).where(User.email == "admin@example.com")).first()
    if existing:
        sync_faculty_program_structure(session)
        sync_me_specification(session)
        return

    users = [
        User(name="System Admin", email="admin@example.com", password_hash=hash_password("password"), role=Role.SUPER_ADMIN),
        User(name="Faculty Dean", email="dean@example.com", password_hash=hash_password("password"), role=Role.DEAN),
        User(name="Program Manager", email="manager@example.com", password_hash=hash_password("password"), role=Role.PROGRAM_MANAGER),
        User(name="Chou Manith", email="teacher@example.com", password_hash=hash_password("password"), role=Role.TEACHER),
        User(name="Pan Chamroeun", email="student@example.com", password_hash=hash_password("password"), role=Role.STUDENT),
    ]
    session.add_all(users)
    session.commit()
    for user in users:
        session.refresh(user)

    sync_faculty_program_structure(session)
    program = session.exec(select(Program).where(Program.code == "ME")).first()
    if not program:
        raise RuntimeError("Mechanical Engineering programme was not created.")

    peos = [PEO(program_id=program.id, code=code, description=description) for code, description in ME_PEOS]
    plos = [
        PLO(program_id=program.id, code=code, description=description)
        for code, description in ME_PLOS
    ]
    session.add_all(peos + plos)
    session.commit()
    for plo in plos:
        session.refresh(plo)

    course = Course(program_id=program.id, code="2SE-ME011", title="Computer Aided Design (Autocad/ Solidwork)", credits=3, curriculum_year=1, curriculum_semester="2")
    session.add(course)
    session.commit()
    session.refresh(course)

    clos = [
        CLO(course_id=course.id, code="CLO1", domain="K,S", description="Apply basic drawing and modification tools in AutoCAD."),
        CLO(course_id=course.id, code="CLO2", domain="A", description="Arrange dimensions, blocks, text, layers, plotting and printing."),
        CLO(course_id=course.id, code="CLO3", domain="S", description="Create isometric drawings and complex 3D objects."),
    ]
    session.add_all(clos)
    session.commit()
    for clo in clos:
        session.refresh(clo)

    mappings = [
        CLOPLOMapping(clo_id=clos[0].id, plo_id=plos[0].id, weight=0.4),
        CLOPLOMapping(clo_id=clos[0].id, plo_id=plos[2].id, weight=0.1),
        CLOPLOMapping(clo_id=clos[1].id, plo_id=plos[4].id, weight=0.25),
        CLOPLOMapping(clo_id=clos[2].id, plo_id=plos[8].id, weight=0.25),
    ]
    assessments = [
        Assessment(clo_id=clos[0].id, name="Quiz I", description="5 questions + drawing exercise", max_score=10, weight=1),
        Assessment(clo_id=clos[0].id, name="Quiz II", description="5 questions + drawing exercise", max_score=10, weight=1),
        Assessment(clo_id=clos[0].id, name="Practice", description="Drawing complex objects", max_score=10, weight=1),
        Assessment(clo_id=clos[1].id, name="Assignment", description="Layer, text and plotting task", max_score=20, weight=1),
        Assessment(clo_id=clos[2].id, name="Final Exam", description="3D modeling task", max_score=50, weight=1),
    ]
    session.add_all(mappings + assessments)
    session.commit()
    for assessment in assessments:
        session.refresh(assessment)

    teacher = Teacher(user_id=users[3].id, staff_no="T-001")
    second_teacher_user = User(name="Sok Dara", email="teacher2@example.com", password_hash=hash_password("password"), role=Role.TEACHER)
    session.add(second_teacher_user)
    session.commit()
    session.refresh(second_teacher_user)
    second_teacher = Teacher(user_id=second_teacher_user.id, staff_no="T-002")
    student = Student(user_id=users[4].id, student_no="241403", name_kh="ប៉ាន់ ចំរើន", name_en="Pan Chamroeun")
    other_students = [
        Student(student_no="241192", name_kh="ប៉ុក បញ្ញា", name_en="Pok Panha"),
        Student(student_no="241910", name_kh="ប្រាក់ ភីលីឧត្ដម", name_en="Prak Phily Oudom"),
        Student(student_no="240052", name_kh="ផូ ប៊ុនណាត", name_en="Pho Bunnat"),
    ]
    session.add_all([teacher, second_teacher, student] + other_students)
    session.commit()
    session.refresh(teacher)
    session.refresh(second_teacher)
    session.refresh(student)
    for item in other_students:
        session.refresh(item)

    course_class = CourseClass(course_id=course.id, academic_year="2025-2026", semester="2", name="20ME CAD")
    session.add(course_class)
    session.commit()
    session.refresh(course_class)

    session.add(ClassTeacher(class_id=course_class.id, teacher_id=teacher.id))
    for item in [student] + other_students:
        session.add(ClassStudent(class_id=course_class.id, student_id=item.id))

    sample_scores = {
        "241403": [8, 7, 7, 16, 42],
        "241192": [5, 5, 4, 14, 35],
        "241910": [0, 5, 3, 12, 28],
        "240052": [8, 8, 7, 17, 43],
    }
    students = [student] + other_students
    for current_student in students:
        for assessment, score in zip(assessments, sample_scores[current_student.student_no], strict=True):
            session.add(StudentScore(assessment_id=assessment.id, student_id=current_student.id, score=score))

    session.commit()

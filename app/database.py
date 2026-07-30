from collections.abc import Generator

from sqlmodel import Session, SQLModel, create_engine
from sqlalchemy import text

from app.config import settings


connect_args = {"check_same_thread": False} if settings.database_url.startswith("sqlite") else {}
engine_options = {"pool_pre_ping": True} if not settings.database_url.startswith("sqlite") else {}
engine = create_engine(settings.database_url, connect_args=connect_args, **engine_options)


def create_db_and_tables() -> None:
    SQLModel.metadata.create_all(engine)
    ensure_lightweight_migrations()


def ensure_lightweight_migrations() -> None:
    if not settings.database_url.startswith("sqlite"):
        return
    with engine.begin() as connection:
        course_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(course)"))}
        if "curriculum_year" not in course_columns:
            connection.execute(text("ALTER TABLE course ADD COLUMN curriculum_year INTEGER"))
        if "curriculum_semester" not in course_columns:
            connection.execute(text("ALTER TABLE course ADD COLUMN curriculum_semester VARCHAR"))
        if "cohort_id" not in course_columns:
            connection.execute(text("ALTER TABLE course ADD COLUMN cohort_id INTEGER"))
        class_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(courseclass)"))}
        if "semester_start" not in class_columns:
            connection.execute(text("ALTER TABLE courseclass ADD COLUMN semester_start VARCHAR"))
        if "semester_end" not in class_columns:
            connection.execute(text("ALTER TABLE courseclass ADD COLUMN semester_end VARCHAR"))
        user_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(user)"))}
        if "faculty_id" not in user_columns:
            connection.execute(text("ALTER TABLE user ADD COLUMN faculty_id INTEGER"))
        if "program_id" not in user_columns:
            connection.execute(text("ALTER TABLE user ADD COLUMN program_id INTEGER"))
        if "avatar_name" not in user_columns:
            connection.execute(
                text("ALTER TABLE user ADD COLUMN avatar_name VARCHAR DEFAULT '' NOT NULL")
            )
        student_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(student)"))}
        if "user_id" in student_columns:
            duplicate_student_account = connection.execute(text(
                "SELECT user_id FROM student WHERE user_id IS NOT NULL "
                "GROUP BY user_id HAVING COUNT(*) > 1 LIMIT 1"
            )).first()
            if duplicate_student_account is None:
                connection.execute(text(
                    "CREATE UNIQUE INDEX IF NOT EXISTS uq_student_user_account "
                    "ON student(user_id) WHERE user_id IS NOT NULL"
                ))
        plo_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(plo)"))}
        if plo_columns and "plo_version_id" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN plo_version_id INTEGER"))
        if plo_columns and "domain" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN domain VARCHAR DEFAULT 'Knowledge'"))
        if plo_columns and "bloom_level" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN bloom_level VARCHAR DEFAULT 'C1'"))
        if plo_columns and "status" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN status VARCHAR DEFAULT 'Active'"))
        if plo_columns and "remark" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN remark VARCHAR DEFAULT ''"))
        if plo_columns and "created_by" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN created_by INTEGER"))
        if plo_columns and "created_at" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN created_at DATETIME"))
            connection.execute(text("UPDATE plo SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL"))
        if plo_columns and "updated_at" not in plo_columns:
            connection.execute(text("ALTER TABLE plo ADD COLUMN updated_at DATETIME"))
        peo_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(peo)"))}
        if peo_columns and "plo_version_id" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN plo_version_id INTEGER"))
        if peo_columns and "status" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN status VARCHAR DEFAULT 'Active'"))
        if peo_columns and "remark" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN remark VARCHAR DEFAULT ''"))
        if peo_columns and "created_by" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN created_by INTEGER"))
        if peo_columns and "created_at" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN created_at DATETIME"))
            connection.execute(text("UPDATE peo SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL"))
        if peo_columns and "updated_at" not in peo_columns:
            connection.execute(text("ALTER TABLE peo ADD COLUMN updated_at DATETIME"))
        peo_mapping_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(peoplomapping)"))}
        if peo_mapping_columns and "weight" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN weight FLOAT DEFAULT 0"))
            connection.execute(text("UPDATE peoplomapping SET weight = 100 WHERE weight IS NULL OR weight = 0"))
        if peo_mapping_columns and "program_id" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN program_id INTEGER"))
        if peo_mapping_columns and "plo_version_id" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN plo_version_id INTEGER"))
        if peo_mapping_columns and "mapping_mode" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN mapping_mode VARCHAR DEFAULT 'percentage'"))
        if peo_mapping_columns and "is_mapped" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN is_mapped BOOLEAN DEFAULT 1"))
        if peo_mapping_columns and "contribution_percentage" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN contribution_percentage FLOAT DEFAULT 0"))
            connection.execute(text("UPDATE peoplomapping SET contribution_percentage = weight WHERE contribution_percentage IS NULL OR contribution_percentage = 0"))
        if peo_mapping_columns and "created_by" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN created_by INTEGER"))
        if peo_mapping_columns and "created_at" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN created_at DATETIME"))
            connection.execute(text("UPDATE peoplomapping SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL"))
        if peo_mapping_columns and "updated_at" not in peo_mapping_columns:
            connection.execute(text("ALTER TABLE peoplomapping ADD COLUMN updated_at DATETIME"))
        enrollment_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(classstudent)"))}
        if enrollment_columns and "status" not in enrollment_columns:
            connection.execute(text("ALTER TABLE classstudent ADD COLUMN status VARCHAR DEFAULT 'Active'"))
        if enrollment_columns and "promoted_to_class_id" not in enrollment_columns:
            connection.execute(text("ALTER TABLE classstudent ADD COLUMN promoted_to_class_id INTEGER"))
        if enrollment_columns and "promoted_at" not in enrollment_columns:
            connection.execute(text("ALTER TABLE classstudent ADD COLUMN promoted_at DATETIME"))
        if enrollment_columns and "promoted_by_user_id" not in enrollment_columns:
            connection.execute(text("ALTER TABLE classstudent ADD COLUMN promoted_by_user_id INTEGER"))
        score_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(studentscore)"))}
        if score_columns and "status" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN status VARCHAR DEFAULT 'Draft'"))
        if score_columns and "locked" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN locked BOOLEAN DEFAULT 0"))
        if score_columns and "updated_at" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN updated_at DATETIME"))
        if score_columns and "submitted_at" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN submitted_at DATETIME"))
        if score_columns and "submitted_by_user_id" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN submitted_by_user_id INTEGER"))
        if score_columns and "entered_by_user_id" not in score_columns:
            connection.execute(text("ALTER TABLE studentscore ADD COLUMN entered_by_user_id INTEGER"))

        # Dedicated teacher-owned score tables are created by SQLModel.metadata.create_all().
        # These indexes make class score entry and reporting fast on larger cohorts.
        connection.execute(text(
            "CREATE INDEX IF NOT EXISTS ix_teacher_score_class_student "
            "ON teacher_assessment_score(course_class_id, student_id)"
        ))
        connection.execute(text(
            "CREATE INDEX IF NOT EXISTS ix_teacher_score_assessment_status "
            "ON teacher_assessment_score(assessment_id, status)"
        ))
        connection.execute(text(
            "CREATE INDEX IF NOT EXISTS ix_teacher_score_period "
            "ON teacher_assessment_score(academic_year, semester)"
        ))


def get_session() -> Generator[Session, None, None]:
    with Session(engine) as session:
        yield session

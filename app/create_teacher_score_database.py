"""Create/update the teacher score database tables safely.

Run from the project root:
    python -m app.create_teacher_score_database
"""
from sqlalchemy import inspect

from app.database import create_db_and_tables, engine


def main() -> None:
    create_db_and_tables()
    tables = set(inspect(engine).get_table_names())
    required = {"teacher_assessment_score", "teacher_score_submission"}
    missing = required - tables
    if missing:
        raise RuntimeError(f"Database update failed; missing tables: {sorted(missing)}")
    print("Teacher score database updated successfully.")
    print("Created/verified: teacher_assessment_score")
    print("Created/verified: teacher_score_submission")


if __name__ == "__main__":
    main()

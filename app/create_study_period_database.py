"""Create and synchronize the global study-period database tables.

Run from project root:
    python -m app.create_study_period_database
"""
from sqlmodel import Session

from app.database import create_db_and_tables, engine
from app.main import seed_study_periods


def main() -> None:
    create_db_and_tables()
    with Session(engine) as session:
        seed_study_periods(session)
    print("Global Study Period synchronized from Academic Semester records.")
    print("Created/verified: studyperiod")
    print("Created/verified: userstudyperiodpreference")
    print("Source: Admin > Academic Structure > Semesters")


if __name__ == "__main__":
    main()

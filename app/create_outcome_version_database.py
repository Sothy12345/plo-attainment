"""Run once if you want to initialise outcome-version tables before starting the app."""
from sqlmodel import Session, select
from app.database import create_db_and_tables, engine
from app.models import PEO, PLO, PLOVersion, Program

create_db_and_tables()
with Session(engine) as session:
    for program in session.exec(select(Program)).all():
        versions = session.exec(select(PLOVersion).where(PLOVersion.programme_id == program.id)).all()
        if not versions:
            version = PLOVersion(programme_id=program.id, version_name="V1", status="Active")
            session.add(version); session.flush()
        else:
            version = versions[0]
        for row in session.exec(select(PLO).where(PLO.program_id == program.id, PLO.plo_version_id == None)).all():
            row.plo_version_id = version.id; session.add(row)
        for row in session.exec(select(PEO).where(PEO.program_id == program.id, PEO.plo_version_id == None)).all():
            row.plo_version_id = version.id; session.add(row)
    session.commit()
print("Outcome version database is ready.")

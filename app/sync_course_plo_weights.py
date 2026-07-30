from sqlmodel import Session, select

from app.database import engine
from app.models import CLO, CLOPLOMapping, Course, CoursePLOMapping, PLO


def normalize(weight: float | int | None) -> float:
    value = float(weight or 0)
    return value * 100 if 0 < value <= 1 else value


def main() -> None:
    updated = 0
    with Session(engine) as session:
        for course in session.exec(select(Course)).all():
            plos = session.exec(select(PLO).where(PLO.program_id == course.program_id)).all()
            totals = {plo.id: 0.0 for plo in plos}
            for clo in session.exec(select(CLO).where(CLO.course_id == course.id)).all():
                for mapping in session.exec(select(CLOPLOMapping).where(CLOPLOMapping.clo_id == clo.id)).all():
                    if mapping.plo_id in totals:
                        totals[mapping.plo_id] += normalize(mapping.weight)
            existing = {
                row.plo_id: row
                for row in session.exec(select(CoursePLOMapping).where(CoursePLOMapping.course_id == course.id)).all()
            }
            for plo in plos:
                value = int(round(max(0, min(100, totals.get(plo.id, 0)))))
                row = existing.get(plo.id)
                if row:
                    row.level = value
                    row.symbol = f"{value}%" if value else ""
                    session.add(row)
                else:
                    session.add(CoursePLOMapping(course_id=course.id, plo_id=plo.id, level=value, symbol=f"{value}%" if value else ""))
                updated += 1
        session.commit()
    print(f"Updated {updated} course-PLO mapping rows from CLO mappings.")


if __name__ == "__main__":
    main()

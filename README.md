# OBE PLO Attainment System

Python web application for Outcome Based Education (OBE) PLO/CLO attainment.

## Technology

- FastAPI
- SQLModel / SQLite for development
- Jinja2 templates
- Bootstrap + HTMX
- openpyxl for Excel import/export

## Roles

- Super Admin: controls all data and settings
- Dean: views faculty/program reports
- Program Manager: creates classes, students, PEO/PLO/CLO, assessment, rubric, mapping
- Teacher: inputs marks and views CLO reports for assigned classes
- Student: views their own report

## Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

## Deploy to Vercel

The project includes a Vercel serverless entry point and configuration. See
[`VERCEL_DEPLOYMENT.md`](VERCEL_DEPLOYMENT.md) for the deployment steps,
environment variables, database options, and storage limitations.

## Demo Login

The app seeds these demo users:

| Role | Email | Password |
|---|---|---|
| Super Admin | admin@example.com | password |
| Dean | dean@example.com | password |
| Program Manager | manager@example.com | password |
| Teacher | teacher@example.com | password |
| Student | student@example.com | password |

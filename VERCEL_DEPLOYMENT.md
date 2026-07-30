# Deploying the OBE PLO Attainment System to Vercel

## 1. Push the project to GitHub

Keep `api/index.py`, `vercel.json`, `requirements.txt`, `app/`, and the database
file at the project root. Do not commit `.env`, exported browser-cookie files,
virtual environments, or database backups.

## 2. Import the repository in Vercel

In Vercel, choose **Add New → Project**, import the repository, leave the
framework preset as **Other**, and deploy from the repository root.

## 3. Add environment variables

In **Project Settings → Environment Variables**, add:

- `SECRET_KEY`: a long random value. You can generate one with
  `python -c "import secrets; print(secrets.token_urlsafe(48))"`.
- `DATABASE_URL`: a hosted PostgreSQL connection URL from Vercel Postgres,
  Neon, or Supabase. This is required for permanent production data.

Redeploy after adding or changing variables.

## Database behavior

- With `DATABASE_URL`, the app creates its tables and initial demo records in
  PostgreSQL. Data remains available across deployments.
- Without `DATABASE_URL`, the app copies the included SQLite database to
  Vercel's `/tmp` directory. This makes previews and demonstrations work, but
  changes can disappear whenever a serverless instance restarts.

## Uploaded documents and profile pictures

Vercel functions do not provide permanent local file storage. The app uses
`/tmp` on Vercel so uploads do not crash, but uploaded documents and avatars
are temporary. For permanent uploads, connect an object-storage service such as
Vercel Blob, Amazon S3, or Cloudflare R2 in a later step.

## Local production check

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app
```

Then open `http://127.0.0.1:8000`.

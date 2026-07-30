import os
import shutil
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


PROJECT_ROOT = Path(__file__).resolve().parent.parent
IS_VERCEL = bool(os.getenv("VERCEL"))


def default_database_url() -> str:
    """Use the bundled database locally and a writable copy on Vercel."""
    configured = os.getenv("DATABASE_URL")
    if configured:
        # Some providers still return the deprecated SQLAlchemy scheme.
        if configured.startswith("postgres://"):
            configured = "postgresql+psycopg://" + configured[len("postgres://"):]
        elif configured.startswith("postgresql://"):
            configured = "postgresql+psycopg://" + configured[len("postgresql://"):]
        return configured

    bundled = PROJECT_ROOT / "obe_plo.db"
    if IS_VERCEL:
        runtime_db = Path("/tmp/obe_plo.db")
        if not runtime_db.exists() and bundled.exists():
            shutil.copy2(bundled, runtime_db)
        return f"sqlite:///{runtime_db}"
    return f"sqlite:///{bundled}"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    app_name: str = "OBE PLO Attainment"
    database_url: str = default_database_url()
    secret_key: str = "change-this-secret-key"
    upload_dir: str = str(Path("/tmp/uploads") if IS_VERCEL else PROJECT_ROOT / "uploads")
    export_dir: str = str(Path("/tmp/exports") if IS_VERCEL else PROJECT_ROOT / "exports")


settings = Settings()

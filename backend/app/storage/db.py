from __future__ import annotations

import sqlite3
from typing import Any

from app.config import settings


DDL = """
CREATE TABLE IF NOT EXISTS jobs (
    job_id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    style TEXT NOT NULL,
    template_id TEXT NOT NULL DEFAULT 'executive_clean',
    status TEXT NOT NULL,
    outline_json TEXT NOT NULL,
    slides_json TEXT NOT NULL,
    parsed_json TEXT NOT NULL DEFAULT '{}',
    material_text TEXT NOT NULL DEFAULT '',
    pptx_url TEXT,
    created_at TEXT NOT NULL
);
"""


def get_conn() -> sqlite3.Connection:
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(settings.database_path)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    rows = conn.execute("PRAGMA table_info(%s)" % table).fetchall()
    cols = [r[1] for r in rows]
    if column not in cols:
        conn.execute(ddl)


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(DDL)
        _ensure_column(
            conn,
            "jobs",
            "template_id",
            "ALTER TABLE jobs ADD COLUMN template_id TEXT NOT NULL DEFAULT 'executive_clean'",
        )
        _ensure_column(conn, "jobs", "material_text", "ALTER TABLE jobs ADD COLUMN material_text TEXT NOT NULL DEFAULT ''")
        _ensure_column(conn, "jobs", "parsed_json", "ALTER TABLE jobs ADD COLUMN parsed_json TEXT NOT NULL DEFAULT '{}'" )
        conn.commit()


def upsert_job(row: dict[str, Any]) -> None:
    sql = """
    INSERT INTO jobs (job_id, title, style, template_id, status, outline_json, slides_json, parsed_json, material_text, pptx_url, created_at)
    VALUES (:job_id, :title, :style, :template_id, :status, :outline_json, :slides_json, :parsed_json, :material_text, :pptx_url, :created_at)
    ON CONFLICT(job_id) DO UPDATE SET
        title=excluded.title,
        style=excluded.style,
        template_id=excluded.template_id,
        status=excluded.status,
        outline_json=excluded.outline_json,
        slides_json=excluded.slides_json,
        parsed_json=excluded.parsed_json,
        material_text=excluded.material_text,
        pptx_url=excluded.pptx_url,
        created_at=excluded.created_at;
    """
    with get_conn() as conn:
        conn.execute(sql, row)
        conn.commit()


def get_job(job_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute("SELECT * FROM jobs WHERE job_id = ?", (job_id,)).fetchone()


def list_jobs(limit: int = 50) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM jobs ORDER BY datetime(created_at) DESC LIMIT ?", (limit,)
        ).fetchall()

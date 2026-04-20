"""SQLite persistence for parsed resumes."""
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from config import DB_PATH


def _conn() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS resumes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                original_filename TEXT NOT NULL,
                participant_name TEXT NOT NULL DEFAULT '',
                extracted_json TEXT NOT NULL,
                edited_json TEXT,
                last_generated_output TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)


def save_resume(filename: str, extracted: dict) -> int:
    """Insert a new resume record. Returns the row id."""
    now = datetime.now(timezone.utc).isoformat()
    name = extracted.get("contact", {}).get("name", "")
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO resumes (original_filename, participant_name, extracted_json, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (filename, name, json.dumps(extracted), now, now),
        )
        return cur.lastrowid


def update_edited(resume_id: int, edited: dict):
    now = datetime.now(timezone.utc).isoformat()
    name = edited.get("contact", {}).get("name", "")
    with _conn() as c:
        c.execute(
            "UPDATE resumes SET edited_json=?, participant_name=?, updated_at=? WHERE id=?",
            (json.dumps(edited), name, now, resume_id),
        )


def update_output_path(resume_id: int, path: str):
    now = datetime.now(timezone.utc).isoformat()
    with _conn() as c:
        c.execute(
            "UPDATE resumes SET last_generated_output=?, updated_at=? WHERE id=?",
            (path, now, resume_id),
        )


def list_resumes() -> list[dict]:
    with _conn() as c:
        rows = c.execute(
            "SELECT id, original_filename, participant_name, created_at, updated_at "
            "FROM resumes ORDER BY updated_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def load_resume(resume_id: int) -> dict | None:
    with _conn() as c:
        row = c.execute("SELECT * FROM resumes WHERE id=?", (resume_id,)).fetchone()
    if not row:
        return None
    d = dict(row)
    d["extracted_json"] = json.loads(d["extracted_json"])
    if d["edited_json"]:
        d["edited_json"] = json.loads(d["edited_json"])
    return d


def delete_resume(resume_id: int):
    with _conn() as c:
        c.execute("DELETE FROM resumes WHERE id=?", (resume_id,))


# Auto-create table on import
init_db()

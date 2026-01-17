import os
import sqlite3
from typing import Optional, List, Dict, Any

DB_PATH = os.getenv("SQLITE_PATH", "./data/emails.db")


def get_conn():
    os.makedirs("./data", exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _has_column(conn: sqlite3.Connection, table: str, col: str) -> bool:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    rows = cur.fetchall()
    return any(r["name"] == col for r in rows)


def _add_column_if_missing(conn: sqlite3.Connection, table: str, col: str, col_def: str) -> None:
    if _has_column(conn, table, col):
        return
    cur = conn.cursor()
    cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_def}")
    conn.commit()


def init_db(conn: sqlite3.Connection):
    cur = conn.cursor()

    # Emails table (plaintext MVP)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            message_id TEXT PRIMARY KEY,
            folder_id TEXT,
            subject TEXT,
            sender TEXT,
            received_dt TEXT,
            weblink TEXT,
            content TEXT NOT NULL
        )
    """)

    # Add Phase 3.5 columns safely for existing DBs
    _add_column_if_missing(conn, "emails", "ingestion_id", "TEXT")
    _add_column_if_missing(conn, "emails", "ingested_at", "TEXT")

    # Ingestion runs table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ingestions (
            ingestion_id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            label TEXT NOT NULL,
            mode TEXT NOT NULL,
            email_count INTEGER NOT NULL DEFAULT 0
        )
    """)

    # Simple metadata table for index state
    cur.execute("""
        CREATE TABLE IF NOT EXISTS index_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    conn.commit()


def set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    cur = conn.cursor()
    cur.execute(
        "INSERT OR REPLACE INTO index_meta (key, value) VALUES (?, ?)",
        (key, value),
    )
    conn.commit()


def get_meta(conn: sqlite3.Connection, key: str) -> Optional[str]:
    cur = conn.cursor()
    cur.execute("SELECT value FROM index_meta WHERE key = ?", (key,))
    row = cur.fetchone()
    if not row:
        return None
    return row["value"]


def upsert_ingestion(conn: sqlite3.Connection, ingestion_id: str, created_at: str, label: str, mode: str) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        INSERT OR IGNORE INTO ingestions (ingestion_id, created_at, label, mode, email_count)
        VALUES (?, ?, ?, ?, 0)
        """,
        (ingestion_id, created_at, label, mode),
    )
    conn.commit()


def set_ingestion_count(conn: sqlite3.Connection, ingestion_id: str, email_count: int) -> None:
    cur = conn.cursor()
    cur.execute(
        "UPDATE ingestions SET email_count = ? WHERE ingestion_id = ?",
        (int(email_count), ingestion_id),
    )
    conn.commit()


def list_ingestions(conn: sqlite3.Connection, limit: int = 50) -> List[Dict[str, Any]]:
    cur = conn.cursor()
    cur.execute(
        """
        SELECT ingestion_id, created_at, label, mode, email_count
        FROM ingestions
        ORDER BY created_at DESC
        LIMIT ?
        """,
        (int(limit),),
    )
    rows = cur.fetchall()
    return [dict(r) for r in rows]

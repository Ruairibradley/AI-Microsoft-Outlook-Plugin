import os
import sqlite3
from typing import Optional

DB_PATH = os.getenv("SQLITE_PATH", "./data/emails.db")


def get_conn():
    os.makedirs("./data", exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


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

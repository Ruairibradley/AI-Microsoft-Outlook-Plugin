import os
import sqlite3

DB_PATH = os.getenv("SQLITE_PATH", "./data/emails.db")


def get_conn():
    os.makedirs("./data", exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection):
    cur = conn.cursor()
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
    conn.commit()


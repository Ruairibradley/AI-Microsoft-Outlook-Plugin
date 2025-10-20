import sqlite3, os

DB_PATH = os.getenv("SQLITE_PATH", "./data/emails.db")

def get_conn():
    os.makedirs("./data", exist_ok=True)
    return sqlite3.connect(DB_PATH)

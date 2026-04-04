"""
database.py — SQLite layer pentru centralizare recipise
"""

import os
import sqlite3
from contextlib import contextmanager

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, 'data', 'recipise.db')
ATTACHMENTS_DIR = os.path.join(BASE_DIR, 'data', 'attachments')

os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


@contextmanager
def db_session():
    conn = get_db()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with db_session() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id TEXT UNIQUE,
                subject TEXT,
                sender TEXT,
                received_date TEXT,
                body TEXT,
                hg_number TEXT,
                br_number TEXT,
                br_date TEXT,
                attachment_count INTEGER DEFAULT 0,
                synced_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id INTEGER REFERENCES emails(id) ON DELETE CASCADE,
                filename TEXT,
                size INTEGER,
                file_path TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_emails_hg ON emails(hg_number);
            CREATE INDEX IF NOT EXISTS idx_emails_entry ON emails(entry_id);
        """)


def email_exists(entry_id):
    with db_session() as conn:
        row = conn.execute("SELECT 1 FROM emails WHERE entry_id = ?", (entry_id,)).fetchone()
        return row is not None


def insert_email(entry_id, subject, sender, received_date, body,
                 hg_number, br_number, br_date, attachment_count):
    with db_session() as conn:
        cursor = conn.execute("""
            INSERT OR IGNORE INTO emails
                (entry_id, subject, sender, received_date, body,
                 hg_number, br_number, br_date, attachment_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (entry_id, subject, sender, received_date, body,
              hg_number, br_number, br_date, attachment_count))
        return cursor.lastrowid


def insert_attachment(email_id, filename, size, file_path):
    with db_session() as conn:
        conn.execute("""
            INSERT INTO attachments (email_id, filename, size, file_path)
            VALUES (?, ?, ?, ?)
        """, (email_id, filename, size, file_path))


def get_all_emails(hg_filter=None, br_filter=None, sender_filter=None):
    with db_session() as conn:
        query = "SELECT * FROM emails WHERE 1=1"
        params = []
        if hg_filter:
            query += " AND hg_number LIKE ?"
            params.append(f"%{hg_filter}%")
        if br_filter:
            query += " AND br_number LIKE ?"
            params.append(f"%{br_filter}%")
        if sender_filter:
            query += " AND sender LIKE ?"
            params.append(f"%{sender_filter}%")
        query += " ORDER BY received_date DESC"
        rows = conn.execute(query, params).fetchall()
        return [dict(r) for r in rows]


def get_email_by_id(email_id):
    with db_session() as conn:
        row = conn.execute("SELECT * FROM emails WHERE id = ?", (email_id,)).fetchone()
        return dict(row) if row else None


def get_attachments_for_email(email_id):
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM attachments WHERE email_id = ? ORDER BY filename",
            (email_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def get_hg_list():
    with db_session() as conn:
        rows = conn.execute("""
            SELECT hg_number, COUNT(*) as cnt
            FROM emails
            WHERE hg_number IS NOT NULL AND hg_number != ''
            GROUP BY hg_number
            ORDER BY hg_number
        """).fetchall()
        return [dict(r) for r in rows]


def get_stats():
    with db_session() as conn:
        total = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
        with_hg = conn.execute(
            "SELECT COUNT(*) FROM emails WHERE hg_number IS NOT NULL AND hg_number != ''"
        ).fetchone()[0]
        with_br = conn.execute(
            "SELECT COUNT(*) FROM emails WHERE br_number IS NOT NULL AND br_number != ''"
        ).fetchone()[0]
        total_att = conn.execute("SELECT COUNT(*) FROM attachments").fetchone()[0]
        return {
            'total': total,
            'with_hg': with_hg,
            'with_br': with_br,
            'total_attachments': total_att,
        }


def update_email_fields(email_id, hg_number=None, br_number=None, br_date=None):
    with db_session() as conn:
        updates = []
        params = []
        if hg_number is not None:
            updates.append("hg_number = ?")
            params.append(hg_number)
        if br_number is not None:
            updates.append("br_number = ?")
            params.append(br_number)
        if br_date is not None:
            updates.append("br_date = ?")
            params.append(br_date)
        if updates:
            params.append(email_id)
            conn.execute(f"UPDATE emails SET {', '.join(updates)} WHERE id = ?", params)


def delete_email(email_id):
    with db_session() as conn:
        conn.execute("DELETE FROM attachments WHERE email_id = ?", (email_id,))
        conn.execute("DELETE FROM emails WHERE id = ?", (email_id,))


# Initialize DB on import
init_db()

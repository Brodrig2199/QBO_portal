import os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime, timezone

DATABASE_URL = os.environ.get("DATABASE_URL", "")

def _conn():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL no está configurado.")
    return psycopg2.connect(DATABASE_URL, sslmode="require")

def init_db():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS qbo_tokens (
              id INT PRIMARY KEY DEFAULT 1,
              access_token TEXT,
              refresh_token TEXT,
              access_expires_at TIMESTAMPTZ,
              updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            );
            INSERT INTO qbo_tokens (id) VALUES (1)
            ON CONFLICT (id) DO NOTHING;
            """)
        conn.commit()

def get_tokens():
    with _conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT * FROM qbo_tokens WHERE id=1;")
            return cur.fetchone()

def save_tokens(access_token: str, refresh_token: str, access_expires_at):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            UPDATE qbo_tokens
            SET access_token=%s,
                refresh_token=%s,
                access_expires_at=%s,
                updated_at=NOW()
            WHERE id=1;
            """, (access_token, refresh_token, access_expires_at))
        conn.commit()

def is_access_token_valid(access_expires_at, skew_seconds=60):
    if not access_expires_at:
        return False
    now = datetime.now(timezone.utc)
    return (access_expires_at - now).total_seconds() > skew_seconds
def is_access_token_valid(access_expires_at, skew_seconds=60) -> bool:
    """
    True si el access_token aún no expira (con margen skew).
    access_expires_at viene como datetime con tz (TIMESTAMPTZ).
    """
    if not access_expires_at:
        return False
    now = datetime.now(timezone.utc)
    return (access_expires_at - now).total_seconds() > skew_seconds

def get_realm_id():
    row = get_tokens()
    return (row or {}).get("realm_id")

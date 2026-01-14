import os
from datetime import datetime, timezone

import psycopg
from psycopg.rows import dict_row

DATABASE_URL = os.environ.get("DATABASE_URL", "")

def _conn():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL no está configurado.")
    # Render Postgres normalmente requiere SSL
    return psycopg.connect(DATABASE_URL, sslmode="require", row_factory=dict_row)

def init_db():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS qbo_tokens (
              id INT PRIMARY KEY DEFAULT 1,
              realm_id TEXT,
              access_token TEXT,
              refresh_token TEXT,
              access_expires_at TIMESTAMPTZ,
              updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            );
            """)
            # Asegura fila id=1
            cur.execute("""
            INSERT INTO qbo_tokens (id) VALUES (1)
            ON CONFLICT (id) DO NOTHING;
            """)
        conn.commit()

def save_tokens(realm_id: str, access_token: str, refresh_token: str, access_expires_at):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            UPDATE qbo_tokens
            SET realm_id=%s,
                access_token=%s,
                refresh_token=%s,
                access_expires_at=%s,
                updated_at=NOW()
            WHERE id=1;
            """, (realm_id, access_token, refresh_token, access_expires_at))
        conn.commit()

def get_tokens():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM qbo_tokens WHERE id=1;")
            return cur.fetchone()

def is_access_token_valid(access_expires_at, skew_seconds=60) -> bool:
    if not access_expires_at:
        return False
    now = datetime.now(timezone.utc)
    return (access_expires_at - now).total_seconds() > skew_seconds


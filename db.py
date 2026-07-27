"""
db.py — SQLite storage for e-signing tokens
"""

import sqlite3
import secrets
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path(__file__).parent / "signing.db"

LINK_EXPIRY_DAYS = 7


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS signing_tokens (
            token            TEXT PRIMARY KEY,
            proposal_ref     TEXT NOT NULL,
            client_name      TEXT NOT NULL,
            client_email     TEXT NOT NULL,
            project_address  TEXT,
            sender_name      TEXT,
            sender_title     TEXT,
            sender_email     TEXT,
            pdf_path         TEXT NOT NULL,
            signed_pdf_path  TEXT,
            status           TEXT NOT NULL DEFAULT 'pending',
            sig_type         TEXT,
            sig_data         TEXT,
            signer_name      TEXT,
            signer_position  TEXT,
            ip_address       TEXT,
            created_at       TEXT NOT NULL,
            expires_at       TEXT NOT NULL,
            signed_at        TEXT,
            source           TEXT NOT NULL DEFAULT 'proposal'
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS quote_invites (
            token            TEXT PRIMARY KEY,
            client_name      TEXT NOT NULL,
            client_email     TEXT NOT NULL,
            project_address  TEXT,
            note             TEXT,
            status           TEXT NOT NULL DEFAULT 'invited',
            proposal_ref     TEXT,
            created_at       TEXT NOT NULL,
            expires_at       TEXT NOT NULL,
            completed_at     TEXT,
            project_type     TEXT
        )
    """)

    # Lightweight migrations for databases created before these columns
    # existed — CREATE TABLE IF NOT EXISTS above is a no-op on an existing
    # table, so older signing.db files need the columns added explicitly.
    existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(signing_tokens)")}
    if "source" not in existing_cols:
        conn.execute(
            "ALTER TABLE signing_tokens ADD COLUMN source TEXT NOT NULL DEFAULT 'proposal'"
        )

    invite_cols = {row["name"] for row in conn.execute("PRAGMA table_info(quote_invites)")}
    if "project_type" not in invite_cols:
        conn.execute("ALTER TABLE quote_invites ADD COLUMN project_type TEXT")

    conn.commit()
    conn.close()


def create_signing_token(
    proposal_ref,
    client_name,
    client_email,
    project_address,
    pdf_path,
    sender_name=None,
    sender_title=None,
    sender_email=None,
    source="proposal",
) -> dict:
    """
    Create a new signing token row. Returns the token + expiry.

    `source` distinguishes a formally-sent Fee Proposal ("proposal", the
    default) from a client self-service Smart Quote ("quote") — used by
    display_status() to label a signed quote as "Quote Accepted" rather
    than "Signed", since an accepted quote is an indicative acceptance
    pending site review, not a finalised proposal.
    """
    token = secrets.token_urlsafe(32)
    now = datetime.utcnow()
    expires = now + timedelta(days=LINK_EXPIRY_DAYS)

    conn = get_conn()
    conn.execute(
        """
        INSERT INTO signing_tokens
            (token, proposal_ref, client_name, client_email, project_address,
             sender_name, sender_title, sender_email,
             pdf_path, status, created_at, expires_at, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?)
        """,
        (
            token,
            proposal_ref,
            client_name,
            client_email,
            project_address,
            sender_name,
            sender_title,
            sender_email,
            pdf_path,
            now.isoformat(),
            expires.isoformat(),
            source,
        ),
    )
    conn.commit()
    conn.close()

    return {"token": token, "expires_at": expires.isoformat()}


def get_token_row(token) -> dict | None:
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM signing_tokens WHERE token = ?", (token,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def is_token_valid(row: dict) -> tuple[bool, str]:
    """Returns (valid, reason_if_invalid)."""
    if row is None:
        return False, "Signing link not found."
    if row["status"] == "signed":
        return False, "This proposal has already been signed."
    if row["status"] == "void":
        return False, "This signing link has been cancelled."
    expires_at = datetime.fromisoformat(row["expires_at"])
    if datetime.utcnow() > expires_at:
        return False, "This signing link has expired."
    return True, ""


def mark_signed(
    token, sig_type, sig_data, signer_name, signer_position, signed_pdf_path, ip_address
):
    conn = get_conn()
    conn.execute(
        """
        UPDATE signing_tokens
        SET status = 'signed',
            sig_type = ?,
            sig_data = ?,
            signer_name = ?,
            signer_position = ?,
            signed_pdf_path = ?,
            ip_address = ?,
            signed_at = ?
        WHERE token = ?
        """,
        (
            sig_type,
            sig_data,
            signer_name,
            signer_position,
            signed_pdf_path,
            ip_address,
            datetime.utcnow().isoformat(),
            token,
        ),
    )
    conn.commit()
    conn.close()


def void_token(token):
    conn = get_conn()
    conn.execute("UPDATE signing_tokens SET status = 'void' WHERE token = ?", (token,))
    conn.commit()
    conn.close()


def list_recent(limit=50):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM signing_tokens ORDER BY created_at DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_stats() -> dict:
    now = datetime.utcnow().isoformat()
    conn = get_conn()
    row = conn.execute("""
        SELECT
            COUNT(*)                                                          AS total,
            SUM(CASE WHEN status = 'signed'                        THEN 1 ELSE 0 END) AS signed,
            SUM(CASE WHEN status = 'void'                          THEN 1 ELSE 0 END) AS voided,
            SUM(CASE WHEN status = 'pending' AND expires_at > ?    THEN 1 ELSE 0 END) AS pending,
            SUM(CASE WHEN status = 'pending' AND expires_at <= ?   THEN 1 ELSE 0 END) AS expired
        FROM signing_tokens
    """, (now, now)).fetchone()
    conn.close()
    return dict(row)


def display_status(row: dict) -> str:
    """Returns the status to show in the UI, upgrading expired pending rows."""
    if row["status"] == "signed" and row.get("source") == "quote":
        return "quote_accepted"
    if row["status"] != "pending":
        return row["status"]
    if datetime.utcnow().isoformat() > row["expires_at"]:
        return "expired"
    return "pending"


# ── Quote invites (client self-service smart quote) ──────────────────────────

QUOTE_INVITE_EXPIRY_DAYS = 14


def create_quote_invite(
    client_name, client_email, project_address=None, note=None, project_type=None
) -> dict:
    """Create a self-service quote invite. Returns the token + expiry.

    `project_type` lets staff pre-select the project type (e.g. "industrial")
    before sending the link, so the client doesn't have to pick it themselves.
    """
    token = secrets.token_urlsafe(32)
    now = datetime.utcnow()
    expires = now + timedelta(days=QUOTE_INVITE_EXPIRY_DAYS)

    conn = get_conn()
    conn.execute(
        """
        INSERT INTO quote_invites
            (token, client_name, client_email, project_address, note,
             status, created_at, expires_at, project_type)
        VALUES (?, ?, ?, ?, ?, 'invited', ?, ?, ?)
        """,
        (token, client_name, client_email, project_address, note,
         now.isoformat(), expires.isoformat(), project_type),
    )
    conn.commit()
    conn.close()
    return {"token": token, "expires_at": expires.isoformat()}


def get_quote_invite(token) -> dict | None:
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM quote_invites WHERE token = ?", (token,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def is_invite_valid(row: dict) -> tuple[bool, str]:
    """Returns (valid, reason_if_invalid)."""
    if row is None:
        return False, "This quote link was not found."
    if row["status"] == "completed":
        return False, "This quote has already been completed and signed."
    if row["status"] == "void":
        return False, "This quote link has been cancelled."
    expires_at = datetime.fromisoformat(row["expires_at"])
    if datetime.utcnow() > expires_at:
        return False, "This quote link has expired. Please contact us for a new one."
    return True, ""


def mark_invite_completed(token, proposal_ref):
    conn = get_conn()
    conn.execute(
        """
        UPDATE quote_invites
        SET status = 'completed', proposal_ref = ?, completed_at = ?
        WHERE token = ?
        """,
        (proposal_ref, datetime.utcnow().isoformat(), token),
    )
    conn.commit()
    conn.close()


def list_pending_invites(limit=100):
    """Invites still awaiting the client (not completed/void)."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM quote_invites WHERE status = 'invited' ORDER BY created_at DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def invite_display_status(row: dict) -> str:
    if row["status"] != "invited":
        return row["status"]
    if datetime.utcnow().isoformat() > row["expires_at"]:
        return "expired"
    return "invited"

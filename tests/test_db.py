"""
Unit tests for db.py, bypassing Flask entirely.

Uses the `client` fixture purely to get the monkeypatched, isolated
DB_PATH + initialised schema — the actual assertions call db functions
directly.
"""

from datetime import datetime, timedelta

import db


def test_create_and_get_signing_token(client):
    info = db.create_signing_token(
        proposal_ref="REF-1",
        client_name="Jane",
        client_email="jane@example.com",
        project_address="1 Test St",
        pdf_path="/tmp/fake.pdf",
    )
    assert "token" in info and info["token"]

    row = db.get_token_row(info["token"])
    assert row is not None
    assert row["client_name"] == "Jane"
    assert row["status"] == "pending"


def test_get_token_row_missing_returns_none(client):
    assert db.get_token_row("does-not-exist") is None


def test_is_token_valid_pending(client):
    info = db.create_signing_token("REF-2", "Jane", "jane@example.com", "addr", "/tmp/f.pdf")
    row = db.get_token_row(info["token"])
    valid, reason = db.is_token_valid(row)
    assert valid is True
    assert reason == ""


def test_is_token_valid_missing_row(client):
    valid, reason = db.is_token_valid(None)
    assert valid is False
    assert "not found" in reason.lower()


def test_is_token_valid_expired(client):
    info = db.create_signing_token("REF-3", "Jane", "jane@example.com", "addr", "/tmp/f.pdf")
    conn = db.get_conn()
    past = (datetime.utcnow() - timedelta(days=1)).isoformat()
    conn.execute("UPDATE signing_tokens SET expires_at = ? WHERE token = ?", (past, info["token"]))
    conn.commit()
    conn.close()

    row = db.get_token_row(info["token"])
    valid, reason = db.is_token_valid(row)
    assert valid is False
    assert "expired" in reason.lower()


def test_mark_signed_updates_row(client):
    info = db.create_signing_token("REF-4", "Jane", "jane@example.com", "addr", "/tmp/f.pdf")
    db.mark_signed(
        info["token"],
        sig_type="draw",
        sig_data="[drawn signature image]",
        signer_name="Jane Client",
        signer_position="Director",
        signed_pdf_path="/tmp/signed.pdf",
        ip_address="127.0.0.1",
    )
    row = db.get_token_row(info["token"])
    assert row["status"] == "signed"
    assert row["signer_name"] == "Jane Client"
    assert row["signed_at"] is not None

    valid, reason = db.is_token_valid(row)
    assert valid is False
    assert "already been signed" in reason


def test_void_token(client):
    info = db.create_signing_token("REF-5", "Jane", "jane@example.com", "addr", "/tmp/f.pdf")
    db.void_token(info["token"])
    row = db.get_token_row(info["token"])
    assert row["status"] == "void"
    valid, reason = db.is_token_valid(row)
    assert valid is False
    assert "cancelled" in reason


def test_list_recent_and_stats(client):
    db.create_signing_token("REF-6", "A", "a@example.com", "addr", "/tmp/f.pdf")
    info2 = db.create_signing_token("REF-7", "B", "b@example.com", "addr", "/tmp/f.pdf")
    db.mark_signed(info2["token"], "draw", "[drawn]", "B", "", "/tmp/s.pdf", "127.0.0.1")

    rows = db.list_recent(50)
    assert len(rows) == 2

    stats = db.get_stats()
    assert stats["total"] == 2
    assert stats["signed"] == 1
    assert stats["pending"] == 1


def test_display_status_upgrades_expired(client):
    info = db.create_signing_token("REF-8", "A", "a@example.com", "addr", "/tmp/f.pdf")
    row = db.get_token_row(info["token"])
    assert db.display_status(row) == "pending"

    conn = db.get_conn()
    past = (datetime.utcnow() - timedelta(days=1)).isoformat()
    conn.execute("UPDATE signing_tokens SET expires_at = ? WHERE token = ?", (past, info["token"]))
    conn.commit()
    conn.close()

    row = db.get_token_row(info["token"])
    assert db.display_status(row) == "expired"


# Quote invites


def test_create_and_get_quote_invite(client):
    info = db.create_quote_invite("Jane", "jane@example.com", "1 Test St")
    row = db.get_quote_invite(info["token"])
    assert row is not None
    assert row["status"] == "invited"
    assert row["client_email"] == "jane@example.com"


def test_is_invite_valid_states(client):
    info = db.create_quote_invite("Jane", "jane@example.com")

    row = db.get_quote_invite(info["token"])
    valid, _ = db.is_invite_valid(row)
    assert valid is True

    db.mark_invite_completed(info["token"], "CEQ-1")
    row = db.get_quote_invite(info["token"])
    valid, reason = db.is_invite_valid(row)
    assert valid is False
    assert "already been completed" in reason


def test_is_invite_valid_missing(client):
    valid, reason = db.is_invite_valid(None)
    assert valid is False
    assert "not found" in reason.lower()


def test_list_pending_invites_excludes_completed(client):
    info1 = db.create_quote_invite("A", "a@example.com")
    info2 = db.create_quote_invite("B", "b@example.com")
    db.mark_invite_completed(info2["token"], "CEQ-2")

    pending = db.list_pending_invites(50)
    tokens = [r["token"] for r in pending]
    assert info1["token"] in tokens
    assert info2["token"] not in tokens


def test_invite_display_status_expired(client):
    info = db.create_quote_invite("A", "a@example.com")
    conn = db.get_conn()
    past = (datetime.utcnow() - timedelta(days=1)).isoformat()
    conn.execute("UPDATE quote_invites SET expires_at = ? WHERE token = ?", (past, info["token"]))
    conn.commit()
    conn.close()

    row = db.get_quote_invite(info["token"])
    assert db.invite_display_status(row) == "expired"

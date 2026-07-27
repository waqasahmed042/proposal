"""
Shared pytest fixtures.

Every test runs against an isolated SQLite DB and isolated storage/ dirs
(via monkeypatch) so tests never touch the real signing.db or real client
documents. Email sending is stubbed at the lowest level (mailer._send) so
tests never hit the network or require real Microsoft Graph credentials —
each stubbed call is recorded in `sent_emails` so tests can assert on it.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

import app as app_module
import db as db_module
import mailer as mailer_module
import auth as auth_module

# A minimal but valid 100x40 transparent PNG, base64-encoded — stands in for
# a real canvas-drawn signature in tests.
VALID_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAGQAAAAoCAYAAAAIeF9DAAAAJ0lEQVR4nO3BgQAAAADDoPlTX+EA"
    "VQEAAAAAAAAAAAAAAAAAAEC3AT6oAAHTb/8VAAAAAElFTkSuQmCC"
)
VALID_SIG_DATA_URL = f"data:image/png;base64,{VALID_PNG_B64}"


@pytest.fixture
def sent_emails():
    """List that fills up with every 'sent' email during a test."""
    return []


@pytest.fixture
def client(tmp_path, monkeypatch, sent_emails):
    """Flask test client wired to isolated DB + storage + stubbed mailer."""

    def fake_send(
        to_email, subject, html_body, attachment_path=None, attachment_name=None
    ):
        sent_emails.append(
            {
                "to": to_email,
                "subject": subject,
                "html_body": html_body,
                "attachment_path": attachment_path,
                "attachment_name": attachment_name,
            }
        )
        return True, ""

    monkeypatch.setattr(mailer_module, "_send", fake_send)

    # These tests exercise staff routes directly, so run them with staff
    # sign-in disabled (login_required becomes a pass-through). Auth itself
    # is covered separately in test_auth.py with AUTH_ENABLED left on.
    monkeypatch.setattr(auth_module, "AUTH_ENABLED", False)

    test_db_path = tmp_path / "test_signing.db"
    monkeypatch.setattr(db_module, "DB_PATH", test_db_path)

    proposals_dir = tmp_path / "storage" / "proposals"
    signed_dir = tmp_path / "storage" / "signed"
    proposals_dir.mkdir(parents=True)
    signed_dir.mkdir(parents=True)
    monkeypatch.setattr(app_module, "PROPOSALS_DIR", proposals_dir)
    monkeypatch.setattr(app_module, "SIGNED_DIR", signed_dir)

    db_module.init_db()

    app_module.app.config.update(TESTING=True)
    with app_module.app.test_client() as c:
        yield c


@pytest.fixture
def make_invite():
    """Factory: create a real quote_invites row and return its dict."""

    def _make(
        client_name="Test Client",
        client_email="test@example.com",
        project_address="1 Test St",
    ):
        info = db_module.create_quote_invite(client_name, client_email, project_address)
        return {
            "client_name": client_name,
            "client_email": client_email,
            "project_address": project_address,
            **info,
        }

    return _make


MINIMAL_PROPOSAL_FORM = {
    "proposal_ref": "TEST-REF-001",
    "contact_name": "Jane Client",
    "contact_email": "jane@example.com",
    "project_address": "1 Test St, Brisbane QLD",
    "company_name": "Client Co",
    "sender_name": "Concept Engineers Team",
    "sender_title": "",
    "sender_email": "admin@conceptengineers.com.au ",
    "is_da": "on",
}

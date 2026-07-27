"""
Tests for POST /api/smart-quote-sign — the client self-service quote
builder's final "sign & submit" step.

These go through the REAL document pipeline (docxtpl -> docx_to_pdf ->
signing.stamp_signature), so they require a working DOCX->PDF converter
on the machine running the tests (MS Word via docx2pdf on Windows/macOS,
or LibreOffice's `soffice` on PATH on Linux — same requirement as the
app itself; see README).
"""

import db
from tests.conftest import VALID_SIG_DATA_URL

BASE_PAYLOAD = {
    "scope": "da",
    "size": "medium",
    "extras": [],
    "tier": "mid",
    "fee_low": 5000,
    "fee_high": 6000,
    "fee_mid": 5500,
    "client_name": "Jane Client",
    "client_email": "jane@example.com",
    "project_address": "1 Test St, Brisbane QLD",
}


def _payload(invite_token, **overrides):
    p = {**BASE_PAYLOAD, "invite_token": invite_token}
    p.update(overrides)
    return p


def test_sign_without_invite_token_rejected(client):
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload("", sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r.status_code == 400
    assert r.get_json().get("error")


def test_sign_with_unknown_invite_token_rejected(client):
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload("not-a-real-token", sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r.status_code == 400


def test_sign_missing_client_name_rejected(client, make_invite):
    invite = make_invite()
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], client_name="", sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r.status_code == 400


def test_sign_invalid_sig_type_rejected(client, make_invite):
    invite = make_invite()
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], sig_type="not-a-real-type", sig_data="x"),
    )
    assert r.status_code == 400


def test_sign_with_drawn_signature_succeeds(client, make_invite, sent_emails):
    invite = make_invite(client_email="jane@example.com")
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r.status_code == 200
    data = r.get_json()
    assert data["success"] is True
    assert data["ref"].startswith("CEQ-")

    # Invite is closed out so the link can't be reused
    invite_row = db.get_quote_invite(invite["token"])
    assert invite_row["status"] == "completed"
    assert invite_row["proposal_ref"] == data["ref"]

    # A signing_tokens row exists and is marked signed
    rows = db.list_recent(50)
    matching = [row for row in rows if row["proposal_ref"] == data["ref"]]
    assert len(matching) == 1
    assert matching[0]["status"] == "signed"

    # Team notification + client signed-copy emails were both sent
    assert len(sent_emails) == 2
    recipients = {e["to"] for e in sent_emails}
    assert "jane@example.com" in recipients


def test_sign_with_typed_signature_succeeds(client, make_invite):
    """
    Regression test: the smart-quote-sign endpoint previously used a
    broken PyMuPDF font name for typed signatures ("helv-oblique"), which
    raised an exception on every submission where the client typed their
    name instead of drawing it. Must succeed end-to-end.
    """
    invite = make_invite()
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], sig_type="type", sig_data="Jane Client"),
    )
    assert r.status_code == 200
    assert r.get_json()["success"] is True


def test_reusing_completed_invite_rejected(client, make_invite):
    invite = make_invite()
    r1 = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r1.status_code == 200

    r2 = client.post(
        "/api/smart-quote-sign",
        json=_payload(invite["token"], sig_type="draw", sig_data=VALID_SIG_DATA_URL),
    )
    assert r2.status_code == 400
    assert "already been completed" in r2.get_json()["error"]


def test_stamping_failure_is_not_reported_as_success(client, make_invite, sent_emails):
    """
    Regression test: a genuine signature-stamping failure was previously
    swallowed silently — the endpoint fell back to the unsigned PDF but
    still returned success=True, marked the DB row 'signed', and emailed
    the client an unsigned document captioned as their signed copy.

    A failure must now surface as an error, and the invite must remain
    usable (not marked completed) so the client can retry.
    """
    invite = make_invite()
    r = client.post(
        "/api/smart-quote-sign",
        json=_payload(
            invite["token"],
            sig_type="draw",
            sig_data="data:image/png;base64,not-valid-base64-data!!!",
        ),
    )
    assert r.status_code == 500
    assert r.get_json().get("error")

    invite_row = db.get_quote_invite(invite["token"])
    assert invite_row["status"] == "invited"

    # No email should have gone out for a failed signing
    assert len(sent_emails) == 0

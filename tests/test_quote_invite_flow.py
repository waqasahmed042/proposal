"""
Tests for the staff-facing 'send a Smart Quote invite' flow:
POST /api/send-quote-invite -> emailed /quote/<token> link -> resend.
"""

import db


def test_send_invite_missing_name(client):
    r = client.post("/api/send-quote-invite", json={"client_email": "a@example.com"})
    assert r.status_code == 400


def test_send_invite_missing_email(client):
    r = client.post("/api/send-quote-invite", json={"client_name": "Jane"})
    assert r.status_code == 400


def test_send_invite_invalid_email(client):
    r = client.post(
        "/api/send-quote-invite",
        json={"client_name": "Jane", "client_email": "not-an-email"},
    )
    assert r.status_code == 400


def test_send_invite_happy_path(client, sent_emails):
    r = client.post(
        "/api/send-quote-invite",
        json={
            "client_name": "Jane Client",
            "client_email": "jane@example.com",
            "project_address": "1 Test St",
            "note": "Great speaking with you today.",
        },
    )
    assert r.status_code == 200
    data = r.get_json()
    assert data["success"] is True
    assert "/quote/" in data["quoteUrl"]
    assert data["token"]
    assert data["emailSent"] is True

    # A real DB row must exist — this was the regression: the previous
    # implementation never persisted an invite at all.
    row = db.get_quote_invite(data["token"])
    assert row is not None
    assert row["client_email"] == "jane@example.com"
    assert row["status"] == "invited"

    # And a real email was composed and "sent" (stubbed) to the client
    assert len(sent_emails) == 1
    assert sent_emails[0]["to"] == "jane@example.com"
    assert data["token"] in sent_emails[0]["html_body"] or data["quoteUrl"] in sent_emails[0]["html_body"]


def test_quote_link_renders_with_prefilled_client(client, make_invite):
    invite = make_invite(client_name="Prefill Test", client_email="prefill@example.com")
    r = client.get(f"/quote/{invite['token']}")
    assert r.status_code == 200
    assert "Prefill Test" in r.get_data(as_text=True)


def test_quote_link_completed_shows_invalid(client, make_invite):
    invite = make_invite()
    db.mark_invite_completed(invite["token"], "CEQ-TEST")
    r = client.get(f"/quote/{invite['token']}")
    assert r.status_code == 200
    assert "already been completed" in r.get_data(as_text=True)


def test_resend_missing_invite_404(client):
    r = client.post("/api/resend-quote/does-not-exist")
    assert r.status_code == 404


def test_resend_completed_invite_400(client, make_invite):
    invite = make_invite()
    db.mark_invite_completed(invite["token"], "CEQ-TEST")
    r = client.post(f"/api/resend-quote/{invite['token']}")
    assert r.status_code == 400


def test_resend_pending_invite_succeeds(client, make_invite, sent_emails):
    invite = make_invite()
    r = client.post(f"/api/resend-quote/{invite['token']}")
    assert r.status_code == 200
    assert r.get_json()["success"] is True
    assert len(sent_emails) == 1

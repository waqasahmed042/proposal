"""Smoke tests: every page route renders without error."""


def test_home_page(client):
    r = client.get("/")
    assert r.status_code == 200


def test_new_proposal_page(client):
    r = client.get("/new-proposal")
    assert r.status_code == 200


def test_new_proposal_dynamic_page(client):
    r = client.get("/new-proposal-dynamic")
    assert r.status_code == 200


def test_smart_quote_invite_page(client):
    """/smart-quote must be the STAFF invite-sending page, not the client funnel."""
    r = client.get("/smart-quote")
    assert r.status_code == 200
    body = r.get_data(as_text=True)
    assert "sendInvite" in body or "Send" in body


def test_smart_quote_preview_page(client):
    r = client.get("/smart-quote-preview")
    assert r.status_code == 200


def test_proposals_dashboard(client):
    r = client.get("/proposals")
    assert r.status_code == 200


def test_unknown_quote_token_404(client):
    r = client.get("/quote/this-token-does-not-exist")
    assert r.status_code == 404


def test_unknown_sign_token_404(client):
    r = client.get("/sign/this-token-does-not-exist")
    assert r.status_code == 404


def test_unknown_document_token_404(client):
    r = client.get("/document/this-token-does-not-exist")
    assert r.status_code == 404

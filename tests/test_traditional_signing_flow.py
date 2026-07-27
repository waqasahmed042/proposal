"""
Tests for the traditional (staff-initiated) proposal flow:
/api/generate-and-send -> /sign/<token> -> /api/sign/<token> -> /document/<token>.

Like test_smart_quote_sign.py, this exercises the real DOCX->PDF converter.
"""

import db
from tests.conftest import MINIMAL_PROPOSAL_FORM, VALID_SIG_DATA_URL


def _create_signing_link(client, **form_overrides):
    """Post the proposal form to the single-request generate-and-send endpoint."""
    form = {**MINIMAL_PROPOSAL_FORM, **form_overrides}
    return client.post("/api/generate-and-send", data=form)


def test_generate_download_returns_document(client):
    """The /generate download path still returns a real file (not the signing flow)."""
    r = client.post("/generate", data={**MINIMAL_PROPOSAL_FORM, "output_fmt": "docx"})
    assert r.status_code == 200
    assert "wordprocessingml" in r.content_type
    assert len(r.get_data()) > 100


def test_generate_and_send_missing_fields(client):
    r = client.post("/api/generate-and-send", data={"proposal_ref": "X"})
    assert r.status_code == 400


def test_generate_and_send_happy_path(client, sent_emails):
    r = _create_signing_link(client)
    assert r.status_code == 200, r.get_data(as_text=True)
    data = r.get_json()
    assert data["success"] is True
    assert "/sign/" in data["signingUrl"]
    assert data["emailSent"] is True

    row = db.get_token_row(data["token"])
    assert row is not None
    assert row["status"] == "pending"
    assert row["client_email"] == "jane@example.com"

    assert len(sent_emails) == 1
    assert sent_emails[0]["to"] == "jane@example.com"


def test_signing_page_renders_for_valid_token(client):
    link = _create_signing_link(client).get_json()

    r = client.get(f"/sign/{link['token']}")
    assert r.status_code == 200
    assert "Jane Client" in r.get_data(as_text=True)


def test_signing_document_preview_serves_pdf(client):
    link = _create_signing_link(client).get_json()

    r = client.get(f"/sign/{link['token']}/document")
    assert r.status_code == 200
    assert r.content_type == "application/pdf"


def test_api_sign_with_drawn_signature(client, sent_emails):
    link = _create_signing_link(client).get_json()
    sent_emails.clear()  # ignore the signing-link email; focus on the sign step

    r = client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "draw",
            "sig_data": VALID_SIG_DATA_URL,
            "signer_name": "Jane Client",
            "signer_position": "Director",
        },
    )
    assert r.status_code == 200
    data = r.get_json()
    assert data["success"] is True

    row = db.get_token_row(link["token"])
    assert row["status"] == "signed"
    assert row["signer_name"] == "Jane Client"

    # Sender notification + client signed-copy — both emailed
    assert len(sent_emails) == 2


def test_api_sign_with_typed_signature(client):
    """Regression coverage for the same font-name bug via the traditional flow."""
    link = _create_signing_link(client).get_json()

    r = client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "type",
            "sig_data": "Jane Client",
            "signer_name": "Jane Client",
        },
    )
    assert r.status_code == 200
    assert r.get_json()["success"] is True


def test_api_sign_rejects_already_signed(client):
    link = _create_signing_link(client).get_json()

    ok = client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "draw",
            "sig_data": VALID_SIG_DATA_URL,
            "signer_name": "Jane Client",
        },
    )
    assert ok.status_code == 200

    again = client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "draw",
            "sig_data": VALID_SIG_DATA_URL,
            "signer_name": "Jane Client",
        },
    )
    assert again.status_code == 400


def test_api_sign_missing_signature_data(client):
    link = _create_signing_link(client).get_json()

    r = client.post(
        f"/api/sign/{link['token']}", json={"sig_type": "draw", "sig_data": ""}
    )
    assert r.status_code == 400


def test_document_download_pdf_and_docx(client):
    link = _create_signing_link(client).get_json()

    r_pdf = client.get(f"/document/{link['token']}")
    assert r_pdf.status_code == 200
    assert r_pdf.content_type == "application/pdf"

    r_docx = client.get(f"/document/{link['token']}?fmt=docx")
    assert r_docx.status_code == 200
    assert "wordprocessingml" in r_docx.content_type


def test_document_download_serves_signed_after_signing(client):
    link = _create_signing_link(client).get_json()
    client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "draw",
            "sig_data": VALID_SIG_DATA_URL,
            "signer_name": "Jane Client",
        },
    )

    r = client.get(f"/document/{link['token']}?download=1")
    assert r.status_code == 200
    assert "_signed.pdf" in r.headers.get("Content-Disposition", "")


def test_resend_missing_proposal_404(client):
    r = client.post("/api/resend/does-not-exist")
    assert r.status_code == 404


def test_resend_signed_proposal_rejected(client):
    link = _create_signing_link(client).get_json()
    client.post(
        f"/api/sign/{link['token']}",
        json={
            "sig_type": "draw",
            "sig_data": VALID_SIG_DATA_URL,
            "signer_name": "Jane Client",
        },
    )

    r = client.post(f"/api/resend/{link['token']}")
    assert r.status_code == 400


def test_resend_pending_proposal_succeeds(client, sent_emails):
    link = _create_signing_link(client).get_json()
    sent_emails.clear()

    r = client.post(f"/api/resend/{link['token']}")
    assert r.status_code == 200
    assert r.get_json()["success"] is True
    assert len(sent_emails) == 1


def test_proposals_dashboard_lists_created_proposal(client):
    _create_signing_link(client)

    r = client.get("/proposals")
    assert r.status_code == 200
    assert "Jane Client" in r.get_data(as_text=True)

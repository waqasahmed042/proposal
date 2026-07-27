"""
Tests for staff sign-in gating (auth.py).

Unlike the shared `client` fixture (which disables auth so it can drive staff
routes directly), these tests keep AUTH_ENABLED on and pretend sign-in IS
configured, to prove:

  * staff pages redirect anonymous users to the login page
  * staff APIs return 401 JSON (not an HTML login page) to anonymous callers
  * client token routes stay OPEN — a client never has to sign in
  * a signed-in staff session can reach staff pages
"""

import pytest

import app as app_module
import auth as auth_module
import db as db_module


@pytest.fixture
def secured_client(tmp_path, monkeypatch):
    """Test client with staff auth ENABLED and treated as configured."""
    monkeypatch.setattr(auth_module, "AUTH_ENABLED", True)
    # Pretend the Entra credentials are present so login_required fails to the
    # login page / 401 (rather than the 503 "not configured" path).
    monkeypatch.setattr(auth_module, "MS_TENANT_ID", "test-tenant")
    monkeypatch.setattr(auth_module, "MS_CLIENT_ID", "test-client")
    monkeypatch.setattr(auth_module, "MS_CLIENT_SECRET", "test-secret")

    test_db_path = tmp_path / "test_signing.db"
    monkeypatch.setattr(db_module, "DB_PATH", test_db_path)
    db_module.init_db()

    app_module.app.config.update(TESTING=True, SECRET_KEY="test-secret-key")
    with app_module.app.test_client() as c:
        yield c


def _sign_in(client, email="staff@conceptengineers.com.au", name="Staff Member"):
    """Simulate a completed Microsoft sign-in by seeding the session."""
    with client.session_transaction() as sess:
        sess["staff_user"] = {"name": name, "email": email}


# ── Staff pages are gated ────────────────────────────────────────────────────

@pytest.mark.parametrize("path", ["/", "/proposals", "/new-proposal", "/smart-quote"])
def test_staff_pages_redirect_anonymous_to_login(secured_client, path):
    r = secured_client.get(path)
    assert r.status_code == 302
    assert "/auth/login" in r.headers["Location"]


def test_staff_api_returns_401_json_for_anonymous(secured_client):
    r = secured_client.post("/api/send-quote-invite", json={})
    assert r.status_code == 401
    body = r.get_json()
    assert body and body.get("login_url") == "/auth/login"


def test_signed_in_staff_can_reach_dashboard(secured_client):
    _sign_in(secured_client)
    r = secured_client.get("/proposals")
    assert r.status_code == 200
    assert "Sent Proposals" in r.get_data(as_text=True) or "Proposals" in r.get_data(as_text=True)


def test_signed_in_staff_sees_their_email_and_logout(secured_client):
    _sign_in(secured_client, email="jane@conceptengineers.com.au")
    r = secured_client.get("/proposals")
    body = r.get_data(as_text=True)
    assert "jane@conceptengineers.com.au" in body
    assert "/auth/logout" in body


# ── Client token routes are NOT gated ────────────────────────────────────────

def test_client_quote_route_not_blocked_by_auth(secured_client):
    """An unknown token 404s — but must NOT redirect to staff login."""
    r = secured_client.get("/quote/some-unknown-token")
    assert r.status_code == 404  # reached the view, token just doesn't exist


def test_client_sign_route_not_blocked_by_auth(secured_client):
    r = secured_client.get("/sign/some-unknown-token")
    assert r.status_code == 404


def test_client_can_open_their_real_quote_link_without_signin(secured_client):
    invite = db_module.create_quote_invite("Client", "client@example.com", "1 St")
    r = secured_client.get(f"/quote/{invite['token']}")
    assert r.status_code == 200  # client reaches their builder with no sign-in
    assert "sign-in required" not in r.get_data(as_text=True).lower()


# ── Login page ───────────────────────────────────────────────────────────────

def test_login_page_renders(secured_client):
    r = secured_client.get("/auth/login")
    assert r.status_code == 200
    assert "Sign in with Microsoft" in r.get_data(as_text=True)


def test_logout_clears_session(secured_client):
    _sign_in(secured_client)
    secured_client.get("/auth/logout")  # redirects to MS logout
    # After logout the dashboard should bounce back to login.
    r = secured_client.get("/proposals")
    assert r.status_code == 302
    assert "/auth/login" in r.headers["Location"]


def test_unconfigured_auth_blocks_staff_fail_closed(tmp_path, monkeypatch):
    """If auth is enabled but NOT configured, staff routes must fail closed
    (503), never fall through to an unprotected view."""
    monkeypatch.setattr(auth_module, "AUTH_ENABLED", True)
    monkeypatch.setattr(auth_module, "MS_TENANT_ID", "")
    monkeypatch.setattr(auth_module, "MS_CLIENT_ID", "")
    monkeypatch.setattr(auth_module, "MS_CLIENT_SECRET", "")
    monkeypatch.setattr(db_module, "DB_PATH", tmp_path / "t.db")
    db_module.init_db()
    app_module.app.config.update(TESTING=True, SECRET_KEY="k")
    with app_module.app.test_client() as c:
        r = c.get("/proposals")
        assert r.status_code == 503

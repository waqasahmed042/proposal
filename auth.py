"""
auth.py — Microsoft Entra ID (Azure AD) sign-in for STAFF pages.

Security model
--------------
There are two completely separate audiences in this app:

  STAFF  — internal team. Must sign in with their Concept Engineers
           Microsoft 365 account. Protected by @login_required.
           Sees the dashboard, all clients, and can generate/send documents.

  CLIENT — external. Never signs in and has no account. Reaches exactly one
           document through an unguessable link (256-bit token from
           secrets.token_urlsafe(32)). The token *is* the authorisation, and
           it maps to a single proposal row, so a client can only ever see
           their own document — never the dashboard or another client's file.

Because the app registration is SINGLE-TENANT, only accounts inside the
Concept Engineers tenant can sign in at all. STAFF_ALLOWED_EMAILS can
narrow that further to named people.

Fail-closed: if auth isn't configured, staff routes are BLOCKED rather than
left open. Set AUTH_ENABLED=false for local development only.

Azure setup (one-time, reuses the existing app registration used for mail):
  1. Portal → Entra ID → App registrations → (your app) → Authentication
  2. Add a platform → Web → Redirect URI:
         {PUBLIC_BASE_URL}/auth/callback
     e.g. https://proposal.conceptengineers.com.au/auth/callback
     (add http://localhost:5000/auth/callback too if you sign in locally)
  3. API permissions → Microsoft Graph → Delegated → User.Read → Add.
     (Admin consent is not required for User.Read.)

.env:
  SECRET_KEY=<long random string>      # REQUIRED — signs session cookies
  MS_TENANT_ID=...                     # reused from the mail config
  MS_CLIENT_ID=...
  MS_CLIENT_SECRET=...
  STAFF_ALLOWED_EMAILS=a@x.com,b@x.com # optional allowlist; blank = whole tenant
  AUTH_ENABLED=true                    # set false ONLY for local dev
"""

import os
from functools import wraps

import msal
from flask import (
    Blueprint,
    current_app,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

auth_bp = Blueprint("auth", __name__)

MS_TENANT_ID = os.getenv("MS_TENANT_ID", "")
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")

# Tenant-specific authority (NOT /common) — this is what restricts sign-in to
# the Concept Engineers tenant. /common would let any Microsoft account in.
AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"

# Pure sign-in: we only need the user's identity, so no Graph scopes are
# requested here. MSAL adds openid/profile/offline_access itself.
SCOPES: list[str] = []

_raw_allowed = os.getenv("STAFF_ALLOWED_EMAILS", "")
STAFF_ALLOWED_EMAILS = {
    e.strip().lower() for e in _raw_allowed.split(",") if e.strip()
}

AUTH_ENABLED = os.getenv("AUTH_ENABLED", "true").strip().lower() not in (
    "false",
    "0",
    "no",
)

_SESSION_USER_KEY = "staff_user"
_SESSION_FLOW_KEY = "auth_flow"
_SESSION_NEXT_KEY = "auth_next"


def is_configured() -> bool:
    """True when the Entra ID credentials needed for sign-in are present."""
    return bool(MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET)


def _redirect_uri() -> str:
    base = (current_app.config.get("PUBLIC_BASE_URL") or "").rstrip("/")
    return f"{base}/auth/callback"


def _msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=AUTHORITY,
        client_credential=MS_CLIENT_SECRET,
    )


def _email_from_claims(claims: dict) -> str:
    """Entra puts the sign-in address in different claims depending on the
    account type — check them in order of reliability."""
    for key in ("preferred_username", "email", "upn", "unique_name"):
        value = claims.get(key)
        if value:
            return str(value).strip().lower()
    return ""


def _is_allowed(email: str) -> bool:
    """Tenant membership is already enforced by the single-tenant authority.
    STAFF_ALLOWED_EMAILS optionally narrows access to named staff."""
    if not STAFF_ALLOWED_EMAILS:
        return True
    return email in STAFF_ALLOWED_EMAILS


def current_user() -> dict | None:
    """The signed-in staff member, or None. Exposed to templates as `user`."""
    if not AUTH_ENABLED:
        return session.get(_SESSION_USER_KEY) or {
            "name": "Developer",
            "email": "dev@localhost",
        }
    return session.get(_SESSION_USER_KEY)


def login_required(view):
    """Protect a STAFF route. Never apply this to client token routes —
    clients have no account and are authorised by their link token."""

    @wraps(view)
    def wrapped(*args, **kwargs):
        if not AUTH_ENABLED:  # local development escape hatch
            return view(*args, **kwargs)

        if session.get(_SESSION_USER_KEY):
            return view(*args, **kwargs)

        # Fail closed: never fall through to the view when misconfigured.
        if not is_configured():
            message = (
                "Staff sign-in is not configured on this server. Set "
                "MS_TENANT_ID, MS_CLIENT_ID and MS_CLIENT_SECRET in .env."
            )
            if request.path.startswith("/api/"):
                return jsonify({"error": message}), 503
            return render_template("login.html", config_error=message), 503

        # APIs get JSON so the frontend can show a real message instead of
        # trying to parse a login page as JSON.
        if request.path.startswith("/api/"):
            return jsonify({"error": "Sign-in required.", "login_url": "/auth/login"}), 401

        session[_SESSION_NEXT_KEY] = request.full_path if request.query_string else request.path
        return redirect(url_for("auth.login"))

    return wrapped


@auth_bp.route("/auth/login")
def login():
    """Show the sign-in page (or bounce straight to Microsoft)."""
    if current_user():
        return redirect("/")
    if not is_configured():
        return render_template(
            "login.html",
            config_error=(
                "Staff sign-in is not configured on this server. Set "
                "MS_TENANT_ID, MS_CLIENT_ID and MS_CLIENT_SECRET in .env."
            ),
        ), 503
    return render_template("login.html", error=request.args.get("error"))


@auth_bp.route("/auth/start")
def start():
    """Kick off the Microsoft authorization-code flow."""
    if not is_configured():
        return redirect(url_for("auth.login"))

    # initiate_auth_code_flow generates and stores state + nonce + PKCE
    # verifier; acquire_token_by_auth_code_flow validates them on return,
    # which is what protects this flow against CSRF / replay.
    flow = _msal_app().initiate_auth_code_flow(SCOPES, redirect_uri=_redirect_uri())
    session[_SESSION_FLOW_KEY] = flow
    return redirect(flow["auth_uri"])


@auth_bp.route("/auth/callback")
def callback():
    """Microsoft redirects back here with the authorization code."""
    flow = session.pop(_SESSION_FLOW_KEY, None)
    if not flow:
        # No flow in session — stale link, or the cookie was dropped.
        return redirect(url_for("auth.login", error="Sign-in session expired. Please try again."))

    try:
        result = _msal_app().acquire_token_by_auth_code_flow(flow, request.args)
    except ValueError as exc:
        # Raised on state mismatch / malformed response.
        print(f"[auth] auth code flow rejected: {exc}")
        return redirect(url_for("auth.login", error="Sign-in failed. Please try again."))

    if "error" in result:
        detail = result.get("error_description") or result.get("error")
        print(f"[auth] token error: {detail}")
        return redirect(url_for("auth.login", error="Sign-in failed. Please try again."))

    claims = result.get("id_token_claims") or {}
    email = _email_from_claims(claims)
    name = claims.get("name") or email or "Staff"

    if not email:
        return redirect(url_for("auth.login", error="Could not read your account email."))

    if not _is_allowed(email):
        print(f"[auth] blocked sign-in for {email} (not in STAFF_ALLOWED_EMAILS)")
        return redirect(
            url_for("auth.login", error="Your account is not authorised for this system.")
        )

    # Store only identity — no access/refresh tokens are kept in the session.
    session[_SESSION_USER_KEY] = {"name": name, "email": email}
    session.permanent = True
    print(f"[auth] ✅ signed in: {email}")

    next_url = session.pop(_SESSION_NEXT_KEY, None)
    # Only allow internal redirects — never bounce to an attacker-supplied host.
    if not next_url or not next_url.startswith("/") or next_url.startswith("//"):
        next_url = "/"
    return redirect(next_url)


@auth_bp.route("/auth/logout")
def logout():
    """Clear the local session and sign out of Microsoft."""
    session.pop(_SESSION_USER_KEY, None)
    session.clear()

    base = (current_app.config.get("PUBLIC_BASE_URL") or "").rstrip("/")
    if is_configured():
        return redirect(
            f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={base}/auth/login"
        )
    return redirect(url_for("auth.login"))

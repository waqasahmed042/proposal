# Concept Engineers — Fee Proposal Automation
## Project Worksheet — What Was Built, Complete Workflow, and Library Usage

Source: https://github.com/waqasahmed042/proposal (branch: `main`)

---

## 1. What This Project Is

A Flask web app for a civil engineering firm (Concept Engineers) that:
- Generates fee proposals from a Word template
- Sends them to clients for e-signature (self-hosted — no DocuSeal, no third-party e-sign service)
- Lets clients build their own quote via a self-service link ("Smart Quote") and sign it inline
- Tracks every proposal/invite in a dashboard with status (pending/signed/expired/void)
- Gates all internal/staff pages behind Microsoft Entra ID (Azure AD) sign-in, while every client-facing page is authorized only by an unguessable link token (no login)

---

## 2. Repository Structure

```
app.py                          Flask routes, PDF conversion, quote-sign endpoint
auth.py                         Microsoft Entra ID (Azure AD) staff sign-in
db.py                           SQLite schema + queries
mailer.py                       Email sending via Microsoft Graph API
signing.py                      PDF signature stamping (PyMuPDF)
build_template_from_reference.py  Dev tool: builds master_proposal_template.docx from client's real doc
calibrate_signature_box.py      Dev tool: visually check signature placement
check_template.py               Dev tool: inspect docxtpl tags in the template
master_proposal_template.docx   The Word template rendered for every proposal
requirements.txt                Python dependencies
.env.example                    Environment variable reference
pytest.ini                      Test configuration

templates/
  home.html                    Landing page (3 entry points)
  login.html                   Staff sign-in page
  form.html                    Staff-facing "New Proposal" form
  proposals.html                Sent-proposals dashboard
  smart_quote_invite.html       Staff-facing "send a Smart Quote link" page
  smart_quote.html              Client-facing Smart Quote builder + inline signing
  new_proposal_dynamic.html     Dynamic variant of the proposal form
  sign.html                     Mobile signing page for traditionally-sent proposals

tests/
  conftest.py                   Shared fixtures (isolated temp DB/storage)
  test_auth.py                  Staff sign-in / access control
  test_db.py                    db.py functions directly
  test_pages.py                 Every route renders; unknown tokens 404
  test_quote_invite_flow.py     Staff sends invite → client link works → resend
  test_signing_module.py        signing.stamp_signature() in isolation
  test_smart_quote_sign.py      Client self-service sign flow end-to-end
  test_traditional_signing_flow.py  Staff-sent proposal → client signs → download
```

---

## 3. Complete Workflow

### A. Access model (two separate audiences)

| Audience | How they get in | Protected by |
|---|---|---|
| **Staff** (internal team) | Sign in with Microsoft 365 work account | `@auth.login_required` on every staff route (`/`, `/proposals`, `/new-proposal`, `/smart-quote`, `/api/generate-and-send`, …). Single-tenant Azure app registration — only accounts inside the Concept Engineers tenant can sign in at all; `STAFF_ALLOWED_EMAILS` can narrow further. **Fails closed**: if auth isn't configured, staff routes return 503 rather than opening up. |
| **Clients** (external) | An unguessable 256-bit link token (`secrets.token_urlsafe(32)`) | The token *is* the authorization — it maps to exactly one row in the database, so a client can only ever reach their own document, never the dashboard or anyone else's file. These routes (`/quote/<token>`, `/sign/<token>`, `/document/<token>`, `/api/sign/<token>`, `/api/smart-quote-sign`) are intentionally **not** behind login. |

### B. Flow 1 — Traditional "New Proposal" (staff-initiated)

1. Staff signs in → visits `/new-proposal` → fills out `form.html`
2. Submits to `/api/generate-and-send`:
   - Renders `master_proposal_template.docx` via **docxtpl** (Jinja templating) with the form's field values (`build_context()`)
   - Converts DOCX → PDF (`docx_to_pdf()` — MS Word via `docx2pdf` on Windows/macOS, LibreOffice headless on Linux/production)
   - Saves the PDF to `storage/proposals/`
   - Creates a `signing_tokens` row in SQLite (`db.create_signing_token()`), expiring in 7 days
   - Emails the client a signing link via **Microsoft Graph API** (`mailer.send_signing_link_email()`)
3. Client opens `/sign/<token>` (`sign.html`) — no login. Draws, types, or uploads a signature.
4. Submits to `/api/sign/<token>`:
   - `signing.stamp_signature()` finds the `{{signature}}` / `{{date}}` placeholders left in the PDF text layer (via **PyMuPDF**) and stamps the signature image (or typed name) and date directly over them
   - `db.mark_signed()` updates the row to `status='signed'`
   - Emails both the client (their signed copy) and the firm (signed notification) — each with the signed PDF attached
5. Staff can track everything on `/proposals` (`proposals.html`) — status, resend, view/download PDF & DOCX.

### C. Flow 2 — Smart Quote (client self-service)

1. Staff visits `/smart-quote` (`smart_quote_invite.html`) → enters client name/email/project address → `/api/send-quote-invite`:
   - Creates a `quote_invites` row (`db.create_quote_invite()`), expiring in 14 days
   - Emails the client a `/quote/<token>` link (`mailer.send_quote_invite_email()`)
2. Client opens `/quote/<token>` → `smart_quote.html` — picks scope (DA/DD/UU/UnityWater), project size, add-ons, and service tier (Bronze/Silver/Gold/Platinum). Price updates live in the browser.
3. Client signs inline (same draw/type/upload widget as Flow 1) and submits to `/api/smart-quote-sign`:
   - Re-validates the invite token server-side
   - Builds the full docxtpl context from the quote selections (`build_template_context()`) — this includes computing fee-table totals for whichever scope sections were selected
   - Renders → converts to PDF → stamps signature (same `signing.py` pipeline as Flow 1) — **in one shot**, since the client is signing at submission time, not via a separate emailed link
   - Marks the invite `completed` (`db.mark_invite_completed()`) so the link can't be reused
   - Emails the signed copy to both client and firm
4. `/smart-quote-preview` exists purely for staff to preview the client-facing wizard UI without a real invite token (submissions from it are rejected server-side).

### D. Document template mechanics (the clever part)

`master_proposal_template.docx` was built from the client's real branded document via `build_template_from_reference.py`. It has:
- **Scope switches** (`is_da`, `is_dd`, `is_uu_minor`, `is_uu_major`, `is_unitywater`) — Jinja conditionals in the DOCX that include/remove entire sections and fee tables depending on what was selected
- **Phase toggles** (`has_construction`, `has_as_constructed`) — turn on/off extra rows inside those fee tables
- **Deferred signing tags**: the template uses `[[signature]]` / `[[date]]` (deliberately *not* valid Jinja syntax) in the signature block. After docxtpl renders everything else, `inject_deferred_signing_tags()` converts those two tags into literal `{{signature}}` / `{{date}}` text — which survives PDF conversion untouched, so `signing.py` can find and stamp over them later once the actual signature/date are known (at signing time, not generation time).

---

## 4. Library Usage — What Does What

| Library | Used for | Where |
|---|---|---|
| **Flask** | Web framework — all routes, request/response handling, sessions | `app.py`, `auth.py` |
| **docxtpl** | Renders `master_proposal_template.docx` as a Jinja template — fills in client/project/fee data and conditionally includes/excludes whole sections via scope switches | `app.py` (`build_context()`, `build_template_context()`, `_render_proposal_to()`) |
| **python-docx** | Low-level DOCX manipulation — specifically to find and rewrite the `[[signature]]`/`[[date]]` deferred tags into `{{signature}}`/`{{date}}` *after* docxtpl has already rendered everything else | `app.py` (`inject_deferred_signing_tags()`) |
| **docx2pdf** | Converts the rendered DOCX to PDF using MS Word — works on Windows/macOS dev machines only | `app.py` (`docx_to_pdf()`) |
| **LibreOffice** (`soffice`, external binary, not a Python package) | Same DOCX→PDF conversion, but headless — this is what actually runs in production on the Linux VPS, since Word isn't available there | `app.py` (`docx_to_pdf()`, Linux fallback path) |
| **PyMuPDF** (`fitz`) | Reads the generated PDF's text layer, finds the `{{signature}}`/`{{date}}` placeholder text, redacts it, and stamps the real signature image (or typed name in an oblique font) and date directly on top | `signing.py` |
| **msal** (Microsoft Authentication Library) | Implements the OAuth2 authorization-code flow for staff sign-in against Microsoft Entra ID (Azure AD) — builds the login URL, exchanges the auth code for tokens, validates state/nonce/PKCE | `auth.py` |
| **requests** | HTTP client used to call the Microsoft Graph API directly — both for getting an OAuth app-only access token (client-credentials flow) and for the actual `sendMail` call | `mailer.py` |
| **sqlite3** (Python standard library) | All persistent storage — two tables: `signing_tokens` (traditional proposals) and `quote_invites` (Smart Quote links) | `db.py` |
| **secrets** (standard library) | Generates the unguessable 256-bit URL-safe tokens that authorize client access to their document | `db.py` |
| **python-dotenv** | Loads `.env` file contents into environment variables for local development | `app.py` |
| **gunicorn** | Production WSGI server that actually runs the Flask app on the VPS (behind PM2) | Deployment only, not imported in application code |
| **pytest** | Test runner for the whole `tests/` suite | Dev/CI only |

### Why Microsoft Graph instead of SMTP

The codebase evolved from a dual SMTP/MS Graph setup (visible as commented-out code at the top of `mailer.py`) to **Graph-only**. Sending is app-only OAuth (client-credentials flow, no user interaction, no stored passwords) — the same Azure app registration is reused for both staff sign-in (`auth.py`, delegated `User.Read` permission) and mail sending (`mailer.py`, application `Mail.Send` permission, requires admin consent).

### Why not DocuSeal

The README notes DocuSeal was evaluated and dropped — its self-hosted community edition blocks template-upload API endpoints behind a paid tier, so signing was built fully custom instead (docxtpl + PyMuPDF, as above).

---

## 5. Testing

Every test runs against an isolated temp SQLite DB and temp `storage/` directory (fixtures in `tests/conftest.py`); `mailer._send` is stubbed so no real emails go out and no Microsoft Graph credentials are needed to run the suite.

| Test file | Covers |
|---|---|
| `test_pages.py` | Every page route renders; unknown tokens 404 |
| `test_auth.py` | Staff sign-in / access control behavior |
| `test_db.py` | `db.py` functions directly (both tables) |
| `test_signing_module.py` | `signing.stamp_signature()` in isolation — draw + typed signatures, missing-placeholder case |
| `test_quote_invite_flow.py` | Staff sends a Smart Quote invite → link renders prefilled → resend |
| `test_smart_quote_sign.py` | Full client self-service sign flow, including regression tests for two specific historical bugs: the typed-signature invalid-font bug, and a silent-stamping-failure bug |
| `test_traditional_signing_flow.py` | Full round trip: staff-sent proposal → client signs → document download |

Two of these (`test_smart_quote_sign.py`, `test_traditional_signing_flow.py`) exercise real DOCX→PDF conversion, so they need either MS Word or LibreOffice actually installed — same requirement as production.

---

## 6. Deployment

- Push to `main`/`master` → GitHub Actions (`.github/workflows/deploy.yml`) rsyncs the repo to a VPS (excluding `.env`, `signing.db`, `storage/`), installs `requirements.txt` into a venv, restarts under **PM2 + gunicorn**.
- `signing.db` and `storage/` (real client data) are never overwritten by deploys — they persist independently on the VPS disk.
- Required GitHub secrets: `VPS_HOST`, `VPS_USER`, `VPS_SSH_KEY`.

---

## 7. Notable Engineering Decisions Worth Knowing

1. **Deferred-tag signing trick**: rather than needing static PDF coordinates for the signature, the template uses inert `[[signature]]`/`[[date]]` markers that get converted to real placeholder text *after* Jinja rendering, so PyMuPDF can locate them by text search regardless of what page they land on or how the surrounding content shifts.
2. **`/api/generate-and-send` replaced a base64 round-trip design**: an earlier version generated the document client-side, base64-encoded it, and POSTed the whole thing back to the server to create the signing link. That was slow and could exceed nginx's request-size limit on large templates. The current version renders and stores the document entirely server-side in one request — the browser never sees the document bytes.
3. **Client vs staff routes are deliberately unauthenticated by design for clients** — the security model relies entirely on token unguessability (256-bit, `secrets.token_urlsafe`) rather than accounts, since clients shouldn't need to create one.
4. **Two known historical bugs are permanently regression-tested**: an invalid PyMuPDF font name (`"helv-oblique"` doesn't exist — must be `"heit"`) that broke every typed signature, and a silent-failure bug where a stamping error used to fall back to emailing the *unsigned* document captioned as "signed." Both now fail loudly instead.

---

*Compiled by fetching the live repository via the GitHub API/raw content — reflects the actual current `main` branch, not any earlier draft.*

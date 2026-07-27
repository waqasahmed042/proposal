# Concept Engineers — Fee Proposal Automation

Flask app that generates civil engineering fee proposals from a Word template, sends them for e-signature, and lets clients build & sign their own quote via a self-service link ("Smart Quote"). Signing is fully self-hosted (no DocuSeal, no third-party e-sign service).

## Access model (two separate audiences)

The app has a hard split between staff and clients — see `auth.py`:

- **Staff** — the internal team. Every staff page/API (`/`, `/proposals`, `/new-proposal`, `/smart-quote`, `/api/generate-and-send`, …) is behind `@auth.login_required` and requires **Microsoft Entra ID (Azure AD) sign-in** with a Concept Engineers work account. The app registration is single-tenant, so only accounts inside the tenant can sign in; `STAFF_ALLOWED_EMAILS` can narrow it further to named people. Auth fails **closed** — if sign-in isn't configured, staff routes are blocked (503), never left open.
- **Clients** — external, never sign in, have no account. A client reaches exactly one document through an unguessable 256-bit link token (`/quote/<token>`, `/sign/<token>`, `/document/<token>`, `/api/sign/<token>`, `/api/smart-quote-sign`). The token *is* the authorisation and maps to a single row, so a client can only ever see their own document — never the dashboard or another client's file. These routes are intentionally NOT behind sign-in.

`AUTH_ENABLED=false` bypasses staff sign-in — **local development only**.

## What it does

- **New Proposal** (`/new-proposal`) — staff fill in a form; the app renders `master_proposal_template.docx` via Jinja (docxtpl), converts it to PDF, and sends the client a signing link. It does not download a copy locally — the generated document is always viewable/downloadable from the `/proposals` dashboard.
- **Sent Proposals** (`/proposals`) — dashboard of every proposal sent for signature: status (pending/signed/expired/void), resend, view/download PDF & DOCX.
- **Smart Quote** (`/smart-quote` internal preview; client link is `/quote/<token>`) — a self-service quote builder. The client picks scope, project size, add-ons and service tier, sees a live price estimate, then signs inline. On submit the app generates the document, stamps the signature, saves it, and emails the signed copy to both the client and the firm.
- **E-signing** (`/sign/<token>`) — mobile-friendly signing page (draw / type / upload signature) for proposals sent the traditional way.

## Architecture

| Concern | How it's done |
|---|---|
| Document generation | `docxtpl` renders `master_proposal_template.docx` — built from the client's real branded proposal (`requirements/26000-FP01-Proposal.docx`) via `build_template_from_reference.py`. Scope switches (`is_da`, `is_dd`, `is_uu_minor`, `is_uu_major`, `is_unitywater`) include/remove whole sections & fee tables; `has_construction`/`has_as_constructed` toggle phase rows inside them. To rebuild after the client updates their reference doc: copy it over `master_proposal_template.docx`, adjust paragraph indices in the script if the doc structure changed, run the script. |
| DOCX → PDF | `docx2pdf` (MS Word) on Windows/macOS; **LibreOffice headless** (`soffice`) on Linux — see `docx_to_pdf()` in `app.py` |
| Signature stamping | `signing.py` (PyMuPDF/`fitz`) finds `{{signature}}` / `{{date}}` placeholders left in the rendered PDF and stamps the signature image/typed name + date directly over them. The template reserves 36pt of `space_after` on the signature paragraph specifically so the stamp has real vertical room (~40pt) instead of being squeezed onto one text line — see `SIG_WIDTH`/`SIG_BOTTOM_OVERHANG` in `signing.py` if this ever needs retuning. |
| Storage | SQLite, `signing.db` — tables `signing_tokens` (traditional signing flow) and `quote_invites` (Smart Quote invite links); generated documents live under `storage/proposals/` and `storage/signed/` |
| Email | Microsoft Graph API (`mailer.py`), app-only OAuth (client-credentials flow) — no SMTP passwords |

### Key files

```
app.py                         Flask routes, PDF conversion, quote-sign endpoint
db.py                          SQLite schema + queries (signing_tokens, quote_invites)
mailer.py                      MS Graph email sending (signing links, quote invites, signed copies)
signing.py                     PDF signature stamping (PyMuPDF)
calibrate_signature_box.py     Dev tool to visually check signature placement on the PDF
check_template.py              Dev tool to inspect docxtpl tags in the master template
master_proposal_template.docx  The Word template rendered for every proposal
templates/
  home.html                   Landing page (3 entry points)
  form.html                   Staff-facing "New Proposal" form
  proposals.html              Sent-proposals dashboard
  smart_quote_invite.html     Staff-facing "send a Smart Quote link" page
  smart_quote.html            Client-facing Smart Quote builder + inline signing
  new_proposal_dynamic.html   Dynamic variant of the proposal form
  sign.html                   Mobile signing page for traditionally-sent proposals
```

## Setup

### 1. Requirements

```bash
pip install -r requirements.txt
```

On Linux (and the production VPS) you also need LibreOffice for DOCX→PDF conversion:

```bash
apt install -y --no-install-recommends libreoffice-writer fonts-liberation
```

### 2. Environment variables

Copy `.env.example` to `.env` and fill in:

```
PUBLIC_BASE_URL=http://localhost:5000     # real domain in production — used to build emailed links

MS_TENANT_ID=...
MS_CLIENT_ID=...
MS_CLIENT_SECRET=...
MAIL_SENDER=admin@conceptengineers.com.au   # must be a real, licensed Microsoft 365 mailbox
MAIL_SENDER_NAME=Concept Engineers
```

**Azure App Registration (one-time):**
1. Azure Portal → Microsoft Entra ID → App registrations → New registration. Supported account types: **single tenant**.
2. API permissions → Add a permission → Microsoft Graph → **Application permissions** → `Mail.Send` → **Grant admin consent** (required, or sending fails with 403).
3. Certificates & secrets → New client secret → copy the secret **value** into `MS_CLIENT_SECRET`.

### 3. Run

```bash
python app.py          # dev server, http://localhost:5000
```

In production the VPS runs it under **gunicorn + PM2** (see `.github/workflows/deploy.yml`).

## Running tests

```bash
pip install -r requirements.txt   # includes pytest
pytest -q
```

Each test runs against an isolated temp SQLite DB and temp `storage/` dir (via fixtures in `tests/conftest.py`), and email sending is stubbed at `mailer._send` — no real DB rows, files, or emails from your real setup are touched, and no Microsoft Graph credentials are required to run the suite.

`tests/test_smart_quote_sign.py` and `tests/test_traditional_signing_flow.py` exercise the real DOCX→PDF conversion, so they need either MS Word (via `docx2pdf`, Windows/macOS) or LibreOffice (`soffice` on PATH, Linux) available — the same requirement the app itself has in production.

| File | Covers |
|---|---|
| `test_pages.py` | Every page route renders; unknown tokens 404 |
| `test_db.py` | `db.py` functions directly (signing tokens + quote invites) |
| `test_signing_module.py` | `signing.stamp_signature()` in isolation — draw + typed signatures, missing placeholder |
| `test_quote_invite_flow.py` | Staff sends a Smart Quote invite → link renders prefilled → resend |
| `test_smart_quote_sign.py` | Client self-service sign flow end-to-end, including regression tests for the typed-signature font bug and the silent-stamping-failure bug |
| `test_traditional_signing_flow.py` | Staff-sent proposal → client signs → document download, full round trip |

## Deployment

Pushing to `master` auto-deploys to the VPS via GitHub Actions (`.github/workflows/deploy.yml`): rsyncs the repo (excluding `.env`, `signing.db`, `storage/`), installs `requirements.txt` into a venv, and restarts the app under PM2/gunicorn.

Required GitHub secrets: `VPS_HOST`, `VPS_USER`, `VPS_SSH_KEY` (base64-encoded private key).

`signing.db` and `storage/` are **never overwritten by deploys** — they hold real client data and persist independently on the VPS disk.

## Notes on routes

- `/smart-quote` — staff-facing page to send a client a Smart Quote invite link (`templates/smart_quote_invite.html`). Creates a `quote_invites` DB row via `db.create_quote_invite()` and emails a `/quote/<token>` link.
- `/smart-quote-preview` — internal-only preview of the client funnel UI (no real invite token, submissions are rejected).
- `templates/quote_funnel.html` was replaced by `templates/smart_quote.html` — if you find references to the old name, they're stale.

## Notes for whoever picks this up next

- Sender identity is deliberately generic: `Concept Engineers Team` / `admin@conceptengineers.com.au ` (not a named staff member) so emails aren't tied to one person leaving.
- Placeholder firm details (ABN, phone) are hardcoded in `app.py`'s `build_context()` — update before real client use.
- DocuSeal was evaluated and dropped: its self-hosted community edition blocks all template-upload API endpoints (Pro-only), which is why signing is fully custom.

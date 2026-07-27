"""
mailer.py — Send emails via Microsoft Graph (app-only OAuth).

No passwords needed. Uses an Azure App Registration with the Application
permission **Mail.Send** (admin consent required).

Azure setup (one-time):
  1. https://portal.azure.com → Microsoft Entra ID → App registrations → New registration.
  2. API permissions → Add a permission → Microsoft Graph → Application permissions
     → Mail.Send → Add. Then click "Grant admin consent".
  3. Certificates & secrets → New client secret → copy the VALUE (not the ID).
  4. Put the values in .env (see below). MAIL_SENDER must be a real, licensed
     Microsoft 365 mailbox the app sends as.

.env:
  MS_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  MS_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  MS_CLIENT_SECRET=your-client-secret-value
  MAIL_SENDER=admin@conceptengineers.com.au
  MAIL_SENDER_NAME=Concept Engineers        (optional display name)

Security tip: by default Mail.Send lets the app send as ANY mailbox in the
tenant. To restrict it to just MAIL_SENDER, an admin can create an
ApplicationAccessPolicy in Exchange Online PowerShell:
  New-ApplicationAccessPolicy -AppId <MS_CLIENT_ID> -PolicyScopeGroupId <mail-enabled security group containing the sender> -AccessRight RestrictAccess
"""

import os
import time
import base64

import requests

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

MS_TENANT_ID = os.getenv("MS_TENANT_ID", "")
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
MAIL_SENDER = os.getenv("MAIL_SENDER", "")
MAIL_SENDER_NAME = os.getenv("MAIL_SENDER_NAME", "Concept Engineers")

# Signature images are served by the app over HTTPS (email clients can't load
# local files, and Gmail strips data: URIs, so absolute hosted URLs are the
# only reliable option). They live in static/signature/ — see _signature().
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:5000").rstrip("/")
_SIG = f"{PUBLIC_BASE_URL}/static/signature"

GRAPH_TOKEN_URL = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
GRAPH_SENDMAIL_URL = f"https://graph.microsoft.com/v1.0/users/{MAIL_SENDER}/sendMail"

if not (MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET and MAIL_SENDER):
    print(
        "⚠️  MS Graph mail not configured — set MS_TENANT_ID, MS_CLIENT_ID, "
        "MS_CLIENT_SECRET and MAIL_SENDER in .env. Emails will fail to send."
    )


def _signature(sender_name: str = "", sender_title: str = "") -> str:
    """
    Zac Lemon's branded email signature — a faithful reproduction of
    'zac lemon signature/index.html' with every image (logo, tagline, the
    orange M/E/W/A contact icons, all social icons, watermark) served over
    HTTPS from static/signature/. Used on ALL client-facing emails.

    The sender_name/sender_title args are kept for call-site compatibility but
    intentionally ignored — the signature is always Zac Lemon's per client
    instruction. All social icons link to the company LinkedIn page, as in the
    original export.
    """
    li = "https://www.linkedin.com/company/conceptengineers-au/"
    return f"""
    <p style="font-family:Arial,sans-serif;font-size:14px;color:#1a202c;margin:18px 0 10px;">Kind regards,</p>
    <table cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;font-family:Arial,sans-serif;">
      <tr>
        <td width="601" valign="top" style="padding:0 5.4pt;">
          <table cellspacing="0" cellpadding="0" border="0" width="607" style="border-collapse:collapse;">
            <tr style="height:51.6pt;">
              <td colspan="4" valign="top" style="border-right:solid #FF5C01 1.5pt;padding:0 5.4pt;height:51.6pt;">
                <p style="margin:0 0 8pt;"><img border="0" width="189" height="64" src="{_SIG}/image002.png" alt="Concept Engineers" style="display:block;"></p>
              </td>
              <td colspan="2" valign="top" style="padding:0 5.4pt;height:51.6pt;">
                <p style="margin:0 0 4pt;font-size:11pt;"><b>Zac Lemon</b></p>
                <p style="margin:0;font-size:11pt;">Director &ndash; Civil &amp; Environmental (NER, RPEQ)</p>
              </td>
              <td rowspan="5" valign="top" style="padding:0 5.4pt;height:51.6pt;">
                <p style="margin:0;"><img border="0" width="81" height="157" src="{_SIG}/image003.jpg" alt="" style="display:block;"></p>
              </td>
            </tr>
            <tr>
              <td colspan="4" rowspan="2" style="border-right:solid #FF5C01 1.5pt;padding:0 5.4pt;">
                <p style="margin:0 0 8pt;"><img border="0" width="75" height="19" src="{_SIG}/image004.jpg" alt="" style="display:block;"></p>
                <p style="margin:0;"><img border="0" width="182" height="18" src="{_SIG}/image005.jpg" alt="Your Vision, Our Expertise" style="display:block;"></p>
              </td>
              <td style="border-right:solid #D9D9D9 1pt;padding:0 5.4pt;">
                <p style="margin:0;"><img border="0" width="20" height="20" src="{_SIG}/image006.png" alt="Phone" style="display:block;"></p>
              </td>
              <td style="padding:0 5.4pt;">
                <p style="margin:0;font-size:11pt;">
                  <a href="tel:+61404483608" style="color:#153d2b;">0404 483 608</a>
                  &nbsp;&nbsp;|&nbsp;&nbsp;
                  <a href="tel:+61735056498" style="color:#153d2b;">(07) 3505 6498</a>
                </p>
              </td>
            </tr>
            <tr>
              <td style="border-right:solid #D9D9D9 1pt;padding:0 5.4pt;">
                <p style="margin:0;"><img border="0" width="20" height="20" src="{_SIG}/image007.png" alt="Email" style="display:block;"></p>
              </td>
              <td style="padding:0 5.4pt;">
                <p style="margin:0;font-size:11pt;">
                  <a href="mailto:zac.lemon@conceptengineers.com.au" style="color:#2f8556;">zac.lemon@conceptengineers.com.au</a>
                </p>
              </td>
            </tr>
            <tr>
              <td rowspan="2" style="padding:0 5.4pt;"><a href="{li}"><img border="0" width="26" height="25" src="{_SIG}/image008.png" alt="LinkedIn" style="display:block;"></a></td>
              <td rowspan="2" style="padding:0 5.4pt;"><a href="{li}"><img border="0" width="26" height="26" src="{_SIG}/image009.png" alt="Facebook" style="display:block;"></a></td>
              <td rowspan="2" style="padding:0 5.4pt;"><a href="{li}"><img border="0" width="26" height="26" src="{_SIG}/image010.png" alt="X" style="display:block;"></a></td>
              <td rowspan="2" style="border-right:solid #FF5C01 1.5pt;padding:0 5.4pt;"><a href="{li}"><img border="0" width="26" height="26" src="{_SIG}/image011.png" alt="Instagram" style="display:block;"></a></td>
              <td style="border-right:solid #D9D9D9 1pt;padding:0 5.4pt;">
                <p style="margin:0;"><img border="0" width="20" height="20" src="{_SIG}/image012.png" alt="Web" style="display:block;"></p>
              </td>
              <td style="padding:0 5.4pt;">
                <p style="margin:0;font-size:11pt;"><a href="https://conceptengineers.com.au/" style="color:#2f8556;">conceptengineers.com.au</a></p>
              </td>
            </tr>
            <tr>
              <td style="border-right:solid #D9D9D9 1pt;padding:0 5.4pt;">
                <p style="margin:0;"><img border="0" width="20" height="20" src="{_SIG}/image013.png" alt="Address" style="display:block;"></p>
              </td>
              <td style="padding:0 5.4pt;">
                <p style="margin:0;font-size:11pt;">Level 4, 111 Boundary Street, West End, QLD 4101</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <p style="font-family:Arial,sans-serif;font-size:8pt;color:#666666;margin-top:12px;line-height:1.5;">
      THIS COMMUNICATION MAY CONTAIN CONFIDENTIAL, PRIVILEGED AND/OR OTHERWISE PROPRIETARY MATERIAL AND IS FOR USE ONLY BY THE INTENDED RECIPIENT.<br>
      If you received this email in error, please contact the sender and delete the email and its attachment from all computers.
    </p>
    """


# Access tokens last ~1 hour; cache one instead of requesting per email.
_token_cache = {"token": None, "expires_at": 0.0}


def _get_access_token(force_refresh: bool = False) -> str:
    """Client-credentials token for Graph. Raises on failure."""
    if (
        not force_refresh
        and _token_cache["token"]
        and time.time() < _token_cache["expires_at"] - 60
    ):
        return _token_cache["token"]

    resp = requests.post(
        GRAPH_TOKEN_URL,
        data={
            "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=20,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Token request failed ({resp.status_code}): {resp.text[:300]}"
        )

    data = resp.json()
    _token_cache["token"] = data["access_token"]
    _token_cache["expires_at"] = time.time() + int(data.get("expires_in", 3600))
    return _token_cache["token"]


def _send(
    to_email: str,
    subject: str,
    html_body: str,
    attachment_path: str | None = None,
    attachment_name: str | None = None,
) -> tuple[bool, str]:
    """Low-level send via MS Graph. Returns (success, error_message)."""
    if not (MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET and MAIL_SENDER):
        return False, (
            "MS Graph mail not configured — set MS_TENANT_ID, MS_CLIENT_ID, "
            "MS_CLIENT_SECRET and MAIL_SENDER in .env"
        )

    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": to_email}}],
        "from": {"emailAddress": {"address": MAIL_SENDER, "name": MAIL_SENDER_NAME}},
    }

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        message["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment_name or os.path.basename(attachment_path),
                "contentType": "application/pdf",
                "contentBytes": content_b64,
            }
        ]

    payload = {"message": message, "saveToSentItems": True}

    try:
        token = _get_access_token()
        resp = requests.post(
            GRAPH_SENDMAIL_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=30,
        )
        # Token may have been revoked/expired early — refresh once and retry.
        if resp.status_code == 401:
            token = _get_access_token(force_refresh=True)
            resp = requests.post(
                GRAPH_SENDMAIL_URL,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=30,
            )

        if resp.status_code == 202:  # Graph returns 202 Accepted on success
            return True, ""
        return False, f"Graph sendMail failed ({resp.status_code}): {resp.text[:300]}"
    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        print(f"[mailer] send failed: {err}")
        return False, err


def send_signing_link_email(
    *,
    client_name: str,
    client_email: str,
    proposal_ref: str,
    project_address: str,
    signing_url: str,
    sender_name: str = "Concept Engineers Team",
    sender_title: str = "",
) -> tuple[bool, str]:
    """Sent to the CLIENT when a signing link is created."""
    subject = f"Fee Proposal {proposal_ref} — Ready for Your Signature"

    html_body = f"""
    <div style="font-family: Arial, sans-serif; max-width: 560px; margin: 0 auto; color:#1a202c;">
      <p>Hi {client_name},</p>
      <p>
        Please find your fee proposal for <strong>{project_address}</strong> ready for review.
      </p>
      <p>
        Click the button below to review the proposal and sign electronically:
      </p>
      <p style="text-align:center; margin: 28px 0;">
        <a href="{signing_url}"
           style="background:#e87722; color:white; text-decoration:none; padding:12px 28px;
                  border-radius:6px; font-weight:600; display:inline-block;">
          Review &amp; Sign Proposal
        </a>
      </p>
      <p style="font-size:13px; color:#718096;">
        Or copy this link into your browser:<br>
        <a href="{signing_url}">{signing_url}</a>
      </p>
      <p style="font-size:13px; color:#718096;">This link will expire in 7 days.</p>
      {_signature(sender_name, sender_title)}
    </div>
    """
    return _send(client_email, subject, html_body)


def send_quote_invite_email(
    *,
    client_name: str,
    client_email: str,
    quote_url: str,
    project_address: str | None = None,
    note: str | None = None,
    sender_name: str = "Concept Engineers Team",
    sender_title: str = "",
) -> tuple[bool, str]:
    """Sent to the CLIENT inviting them to build & sign their own quote."""
    subject = "Build Your Fee Proposal — Concept Engineers"

    address_line = (
        f"<p>Regarding your project at <strong>{project_address}</strong>:</p>"
        if project_address
        else ""
    )
    note_block = (
        f"""<p style="background:#f7f9fc; border-left:3px solid #e87722; padding:10px 14px;
              font-size:14px; color:#4a5568;">{note}</p>"""
        if note
        else ""
    )

    html_body = f"""
    <div style="font-family: Arial, sans-serif; max-width: 560px; margin: 0 auto; color:#1a202c;">
      <p>Hi {client_name},</p>
      {address_line}
      <p>
        We've set up a personalised quote builder for you. Choose the engineering
        services you need, see your fee estimate update live, and sign online —
        all in a few minutes, from any device.
      </p>
      {note_block}
      <p style="text-align:center; margin: 28px 0;">
        <a href="{quote_url}"
           style="background:#e87722; color:white; text-decoration:none; padding:13px 30px;
                  border-radius:6px; font-weight:600; display:inline-block; font-size:15px;">
          Build My Quote
        </a>
      </p>
      <p style="font-size:13px; color:#718096;">
        Or copy this link into your browser:<br>
        <a href="{quote_url}">{quote_url}</a>
      </p>
      <p style="font-size:13px; color:#718096;">This link is personal to you and expires in 14 days.</p>
      {_signature(sender_name, sender_title)}
    </div>
    """
    return _send(client_email, subject, html_body)


def send_client_signed_copy_email(
    *,
    client_name: str,
    client_email: str,
    proposal_ref: str,
    project_address: str,
    signed_pdf_path: str,
    sender_name: str = "Concept Engineers Team",
    sender_title: str = "",
) -> tuple[bool, str]:
    """Sent to the CLIENT after signing — attaches their signed copy."""
    subject = f"Your Signed Proposal {proposal_ref} — Concept Engineers"

    html_body = f"""
    <div style="font-family: Arial, sans-serif; max-width: 560px; margin: 0 auto; color:#1a202c;">
      <p>Hi {client_name},</p>
      <p>
        Thank you — your fee proposal for <strong>{project_address}</strong> has been
        signed successfully. A copy of the signed document is attached to this email
        for your records.
      </p>
      <table style="font-size:14px; border-collapse:collapse; margin: 16px 0;">
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Proposal Ref:</td><td><strong>{proposal_ref}</strong></td></tr>
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Project Address:</td><td>{project_address}</td></tr>
      </table>
      <p>We'll be in touch shortly to confirm the next steps.</p>
      {_signature(sender_name, sender_title)}
    </div>
    """
    return _send(
        client_email,
        subject,
        html_body,
        attachment_path=signed_pdf_path,
        attachment_name=f"{proposal_ref}_signed.pdf",
    )


def send_signed_notification_email(
    *,
    sender_email: str,
    client_name: str,
    client_email: str,
    proposal_ref: str,
    project_address: str,
    signed_pdf_path: str,
    signer_name: str,
    signed_at: str,
) -> tuple[bool, str]:
    """Sent to the SENDER (firm) when the client completes signing. Attaches the signed PDF."""
    subject = f"✅ Proposal {proposal_ref} Signed by {client_name}"

    html_body = f"""
    <div style="font-family: Arial, sans-serif; max-width: 560px; margin: 0 auto; color:#1a202c;">
      <p>Good news — the proposal has been signed.</p>
      <table style="font-size:14px; border-collapse:collapse; margin: 16px 0;">
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Proposal Ref:</td><td><strong>{proposal_ref}</strong></td></tr>
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Project Address:</td><td>{project_address}</td></tr>
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Signed By:</td><td>{signer_name}</td></tr>
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Client Email:</td><td>{client_email}</td></tr>
        <tr><td style="padding:4px 12px 4px 0; color:#718096;">Signed At (UTC):</td><td>{signed_at}</td></tr>
      </table>
      <p>The signed PDF is attached to this email for your records.</p>
    </div>
    """
    attachment_name = f"{proposal_ref}_signed.pdf"
    return _send(
        sender_email,
        subject,
        html_body,
        attachment_path=signed_pdf_path,
        attachment_name=attachment_name,
    )


# """
# mailer.py — Concept Engineers
# Supports two mail providers, switched via MAIL_PROVIDER in .env:

#   MAIL_PROVIDER=smtp      → works with Gmail, Outlook.com, any SMTP server
#   MAIL_PROVIDER=msgraph   → requires an organisational Microsoft 365 mailbox

# .env reference
# # SMTP (use this for free testing with Gmail or Outlook.com)
# MAIL_PROVIDER=smtp
# SMTP_HOST=smtp-mail.outlook.com      # or smtp.gmail.com
# SMTP_PORT=587
# SMTP_USER=youraddress@outlook.com    # or @gmail.com
# SMTP_PASS=your-password              # Gmail: use 16-char App Password
# MAIL_SENDER=youraddress@outlook.com  # must match SMTP_USER for Outlook/Gmail

# # MS Graph (production — requires Microsoft 365 organisational mailbox)─
# # MAIL_PROVIDER=msgraph
# # MS_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
# # MS_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
# # MS_CLIENT_SECRET=your-client-secret
# # MAIL_SENDER=noreply@yourcompany.com   ← must be an M365 mailbox in the tenant
# """

# import os
# import smtplib
# import ssl
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.application import MIMEApplication
# from pathlib import Path

# # Config

# MAIL_PROVIDER = os.getenv("MAIL_PROVIDER", "smtp").lower()  # "smtp" | "msgraph"

# # SMTP
# SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
# SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
# SMTP_USER = os.getenv("SMTP_USER", "admin@conceptengineers.com.au ")
# SMTP_PASS = os.getenv("SMTP_PASS", "dpst pulx pnua lwvi")
# MAIL_SENDER = os.getenv("MAIL_SENDER", SMTP_USER)

# # MS Graph
# MS_TENANT_ID = os.getenv("MS_TENANT_ID", "")
# MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
# MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")

# FIRM_NAME = "Concept Engineers"


# def _signature(sender_name: str, sender_title: str = "") -> str:
#     """Email sign-off that skips empty lines and avoids 'Concept Engineers' twice."""
#     lines = [f"<strong>{sender_name}</strong>"]
#     if sender_title:
#         lines.append(sender_title)
#     if "concept engineers" not in sender_name.lower():
#         lines.append("Concept Engineers")
#     return "<p>Yours faithfully,<br>" + "<br>".join(lines) + "</p>"


# # ─
# # LOW-LEVEL SEND
# # ─


# def _send_smtp(
#     to_email: str,
#     subject: str,
#     html_body: str,
#     attachment_path: str | None = None,
#     attachment_name: str | None = None,
# ) -> tuple[bool, str]:
#     """Send via SMTP STARTTLS — works with Outlook.com and Gmail."""
#     if not SMTP_USER or not SMTP_PASS:
#         return False, "SMTP_USER or SMTP_PASS not set in .env"

#     msg = MIMEMultipart("mixed")
#     msg["From"] = f"{FIRM_NAME} <{MAIL_SENDER}>"
#     msg["To"] = to_email
#     msg["Subject"] = subject

#     msg.attach(MIMEText(html_body, "html"))

#     if attachment_path and Path(attachment_path).exists():
#         with open(attachment_path, "rb") as f:
#             part = MIMEApplication(f.read(), _subtype="pdf")
#         part.add_header(
#             "Content-Disposition",
#             "attachment",
#             filename=attachment_name or Path(attachment_path).name,
#         )
#         msg.attach(part)

#     try:
#         context = ssl.create_default_context()
#         with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as server:
#             server.ehlo()
#             server.starttls(context=context)
#             server.login(SMTP_USER, SMTP_PASS)
#             server.sendmail(MAIL_SENDER, to_email, msg.as_string())
#         return True, ""
#     except smtplib.SMTPAuthenticationError:
#         hint = ""
#         if "gmail" in SMTP_HOST:
#             hint = " — Gmail requires a 16-char App Password (not your real password). Enable it at: Google Account → Security → 2-Step Verification → App Passwords."
#         elif "outlook" in SMTP_HOST:
#             hint = " — Make sure SMTP AUTH is enabled for your Outlook.com account."
#         return False, f"SMTP authentication failed{hint}"
#     except Exception as e:
#         return False, f"SMTP error: {type(e).__name__}: {e}"


# def _send_msgraph(
#     to_email: str,
#     subject: str,
#     html_body: str,
#     attachment_path: str | None = None,
#     attachment_name: str | None = None,
# ) -> tuple[bool, str]:
#     """
#     Send via Microsoft Graph API (app-only, client credentials flow).
#     Requires an organisational Microsoft 365 mailbox as MAIL_SENDER.
#     Personal @outlook.com addresses will NOT work here.
#     """
#     if not all([MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MAIL_SENDER]):
#         return (
#             False,
#             "MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, and MAIL_SENDER must all be set for msgraph provider.",
#         )

#     try:
#         import requests
#     except ImportError:
#         return False, "'requests' package not installed — run: pip install requests"

#     # 1. Get access token
#     token_url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
#     token_resp = requests.post(
#         token_url,
#         data={
#             "grant_type": "client_credentials",
#             "client_id": MS_CLIENT_ID,
#             "client_secret": MS_CLIENT_SECRET,
#             "scope": "https://graph.microsoft.com/.default",
#         },
#         timeout=15,
#     )

#     if not token_resp.ok:
#         return (
#             False,
#             f"Graph token error: {token_resp.status_code} {token_resp.text[:200]}",
#         )

#     access_token = token_resp.json().get("access_token")
#     if not access_token:
#         return False, "Graph token response missing access_token"

#     # 2. Build message payload
#     message: dict = {
#         "subject": subject,
#         "body": {"contentType": "HTML", "content": html_body},
#         "toRecipients": [{"emailAddress": {"address": to_email}}],
#     }

#     if attachment_path and Path(attachment_path).exists():
#         import base64

#         with open(attachment_path, "rb") as f:
#             b64 = base64.b64encode(f.read()).decode()
#         message["attachments"] = [
#             {
#                 "@odata.type": "#microsoft.graph.fileAttachment",
#                 "name": attachment_name or Path(attachment_path).name,
#                 "contentType": "application/pdf",
#                 "contentBytes": b64,
#             }
#         ]

#     # 3. Send
#     send_url = f"https://graph.microsoft.com/v1.0/users/{MAIL_SENDER}/sendMail"
#     send_resp = requests.post(
#         send_url,
#         json={"message": message, "saveToSentItems": False},
#         headers={
#             "Authorization": f"Bearer {access_token}",
#             "Content-Type": "application/json",
#         },
#         timeout=20,
#     )

#     if send_resp.status_code == 202:
#         return True, ""

#     return (
#         False,
#         f"Graph sendMail error: {send_resp.status_code} {send_resp.text[:300]}",
#     )


# def _send(
#     to_email: str,
#     subject: str,
#     html_body: str,
#     attachment_path: str | None = None,
#     attachment_name: str | None = None,
# ) -> tuple[bool, str]:
#     """Route to the correct provider based on MAIL_PROVIDER env var."""
#     if MAIL_PROVIDER == "msgraph":
#         return _send_msgraph(
#             to_email, subject, html_body, attachment_path, attachment_name
#         )
#     return _send_smtp(to_email, subject, html_body, attachment_path, attachment_name)


# # ─
# # EMAIL TEMPLATES
# # ─


# def _base_html(body_content: str) -> str:
#     return f"""
# <!DOCTYPE html>
# <html>
# <head>
# <meta charset="UTF-8">
# <meta name="viewport" content="width=device-width, initial-scale=1.0">
# <style>
#   body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
#           background: #f0f4f8; margin: 0; padding: 24px 0; color: #1a202c; }}
#   .wrap {{ max-width: 580px; margin: 0 auto; background: white;
#            border-radius: 10px; border: 1px solid #e2e8f0; overflow: hidden; }}
#   .header {{ background: #1f3355; padding: 24px 32px; }}
#   .header h1 {{ color: white; font-size: 18px; margin: 0; font-weight: 700; }}
#   .header p  {{ color: #93afd4; font-size: 13px; margin: 4px 0 0; }}
#   .body   {{ padding: 28px 32px; font-size: 14px; line-height: 1.7; color: #2d3748; }}
#   .body p {{ margin: 0 0 14px; }}
#   .btn    {{ display: inline-block; background: #2e75b6; color: white !important;
#              padding: 13px 28px; border-radius: 7px; text-decoration: none;
#              font-weight: 600; font-size: 15px; margin: 8px 0 20px; }}
#   .detail {{ background: #f7f9fc; border: 1px solid #e2e8f0; border-radius: 7px;
#              padding: 14px 18px; margin: 16px 0; font-size: 13px; }}
#   .detail b  {{ color: #1f3355; }}
#   .footer {{ background: #f7f9fc; border-top: 1px solid #e2e8f0;
#              padding: 16px 32px; font-size: 12px; color: #718096; line-height: 1.6; }}
# </style>
# </head>
# <body>
# <div class="wrap">
#   <div class="header">
#     <h1>{FIRM_NAME}</h1>
#     <p>conceptengineers.com.au</p>
#   </div>
#   <div class="body">
#     {body_content}
#   </div>
#   <div class="footer">
#     {FIRM_NAME} &nbsp;|&nbsp; ABN 00 000 000 000 &nbsp;|&nbsp; conceptengineers.com.au<br>
#     This email was sent automatically. Please do not reply directly to this email.
#   </div>
# </div>
# </body>
# </html>"""


# # ─
# # PUBLIC API
# # ─


# def send_signing_link_email(
#     client_name: str,
#     client_email: str,
#     proposal_ref: str,
#     project_address: str,
#     signing_url: str,
#     sender_name: str,
#     sender_title: str,
# ) -> tuple[bool, str]:
#     """Email the client their signing link."""
#     subject = f"Fee Proposal {proposal_ref} — Ready for Your Signature"

#     body = f"""
# <!DOCTYPE html>
# <html>
# <head>
# <meta charset="UTF-8">
# <meta name="viewport" content="width=device-width, initial-scale=1.0">
# <style>
#   body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
#           margin: 0; padding: 0; color: #1a202c; background: #ffffff; }}
#   .wrap {{ max-width: 600px; margin: 0 auto; padding: 20px 24px 28px; font-size: 14px; line-height: 1.7; }}
#   p {{ margin: 0 0 16px; }}
#   .btn-row {{ text-align: right; margin: 22px 0; }}
#   .btn {{ display: inline-block; background: #1f3355; color: #ffffff !important;
#           padding: 11px 22px; text-decoration: none; font-weight: 700; font-size: 14px; }}
#   .muted {{ color: #9aa5b1; font-size: 13px; margin: 0 0 6px; }}
#   .muted a {{ color: #2e75b6; word-break: break-all; }}
# </style>
# </head>
# <body>
# <div class="wrap">
#   <p>Hi {client_name},</p>
#   <p>Please find your fee proposal for <b>{project_address}</b> ready for review.</p>
#   <p>Click the button below to review the proposal and sign electronically:</p>
#   <p style="text-align:center; margin: 28px 0;">
#     <a href="{signing_url}"
#        style="background:#e87722; color:white; text-decoration:none; padding:13px 30px;
#             border-radius:6px; font-weight:600; display:inline-block; font-size:15px;">
#         Review &amp; Sign Proposal
#     </a>
# </p>
#   <p class="muted">Or copy this link into your browser:<br>
#      <a href="{signing_url}">{signing_url}</a></p>
#   <p class="muted">This link will expire in 7 days.</p>
#   <p>Yours faithfully,<br><b>{FIRM_NAME} Team</b></p>
# </div>
# </body>
# </html>"""

#     return _send(client_email, subject, body)


# def send_quote_invite_email(
#     *,
#     client_name: str,
#     client_email: str,
#     quote_url: str,
#     project_address: str | None = None,
#     note: str | None = None,
#     sender_name: str = "Concept Engineers Team",
#     sender_title: str = "",
# ) -> tuple[bool, str]:
#     """Sent to the CLIENT inviting them to build & sign their own quote."""
#     subject = "Build Your Fee Proposal — Concept Engineers"

#     address_line = (
#         f"<p>Regarding your project at <strong>{project_address}</strong>:</p>"
#         if project_address
#         else ""
#     )
#     note_block = (
#         f"""<p style="background:#f7f9fc; border-left:3px solid #e87722; padding:10px 14px;
#               font-size:14px; color:#4a5568;">{note}</p>"""
#         if note
#         else ""
#     )

#     html_body = f"""
#     <div style="font-family: Arial, sans-serif; max-width: 560px; margin: 0 auto; color:#1a202c;">
#       <p>Hi {client_name},</p>
#       {address_line}
#       <p>
#         We've set up a personalised quote builder for you. Choose the engineering
#         services you need, see your fee estimate update live, and sign online —
#         all in a few minutes, from any device.
#       </p>
#       {note_block}
#       <p style="text-align:center; margin: 28px 0;">
#         <a href="{quote_url}"
#            style="background:#e87722; color:white; text-decoration:none; padding:13px 30px;
#                   border-radius:6px; font-weight:600; display:inline-block; font-size:15px;">
#           Build My Quote
#         </a>
#       </p>
#       <p style="font-size:13px; color:#718096;">
#         Or copy this link into your browser:<br>
#         <a href="{quote_url}">{quote_url}</a>
#       </p>
#       <p style="font-size:13px; color:#718096;">This link is personal to you and expires in 14 days.</p>
#       {_signature(sender_name, sender_title)}
#     </div>
#     """
#     return _send(client_email, subject, html_body)


# def send_signed_notification_email(
#     sender_email: str,
#     client_name: str,
#     client_email: str,
#     proposal_ref: str,
#     project_address: str,
#     signed_pdf_path: str,
#     signer_name: str,
#     signed_at: str,
# ) -> tuple[bool, str]:
#     """Notify the firm that a proposal has been signed, with signed PDF attached."""
#     subject = f"✅ Proposal Signed — {proposal_ref} ({client_name})"

#     body = _base_html(f"""
#         <p>Hi,</p>
#         <p>The following fee proposal has been signed by the client.</p>
#         <div class="detail">
#           <b>Proposal Ref:</b>  {proposal_ref}<br>
#           <b>Project:</b>       {project_address}<br>
#           <b>Client:</b>        {client_name} ({client_email})<br>
#           <b>Signed by:</b>     {signer_name}<br>
#           <b>Signed at:</b>     {signed_at}
#         </div>
#         <p>The signed PDF is attached to this email and has also been saved to the
#            <code style="background:#f1f5f9;padding:1px 5px;border-radius:4px;">signed_docs/</code>
#            folder on the server.</p>
#         <p>You can now update the proposal status in HubSpot and commence the project.</p>
#     """)

#     return _send(
#         sender_email,
#         subject,
#         body,
#         attachment_path=signed_pdf_path,
#         attachment_name=f"{proposal_ref}_signed.pdf",
#     )


# def send_client_signed_copy_email(
#     client_name: str,
#     client_email: str,
#     proposal_ref: str,
#     project_address: str,
#     signed_pdf_path: str,
#     sender_name: str = "Concept Engineers Team",
#     sender_title: str = "",
# ) -> tuple[bool, str]:
#     """
#     Send the client their own signed copy of the proposal, PDF attached.
#     Goes only to client_email — no CC/BCC to the firm (the firm is notified
#     separately by send_signed_notification_email).
#     """
#     subject = f"Your Signed Proposal — {proposal_ref} | {FIRM_NAME}"

#     body = _base_html(f"""
#         <p>Hi {client_name},</p>
#         <p>Thank you for signing your fee proposal with {FIRM_NAME}. A copy of your
#            fully signed proposal is attached to this email for your records.</p>
#         <div class="detail">
#           <b>Proposal Ref:</b> {proposal_ref}<br>
#           <b>Project:</b> {project_address}
#         </div>
#         <p>We look forward to working with you on this project. If you have any
#            questions in the meantime, please don't hesitate to reach out.</p>
#         <p>Kind regards,<br><b>{sender_name}</b><br>{sender_title}<br>{FIRM_NAME}</p>
#     """)

#     return _send(
#         client_email,
#         subject,
#         body,
#         attachment_path=signed_pdf_path,
#         attachment_name=f"{proposal_ref}_signed.pdf",
#     )


# def send_reminder_email(
#     client_name: str,
#     client_email: str,
#     proposal_ref: str,
#     project_address: str,
#     signing_url: str,
#     reminder_number: int,
#     sender_name: str,
#     sender_title: str,
# ) -> tuple[bool, str]:
#     """
#     Automated follow-up email. Tone shifts based on reminder_number:
#       1–2 : gentle nudge
#       3–5 : friendly follow-up
#       6+  : easy-out / closing
#     """
#     if reminder_number <= 2:
#         subject = f"Reminder — Fee Proposal {proposal_ref} awaiting your signature"
#         opener = f"Just a friendly reminder that your fee proposal from {FIRM_NAME} is ready for your signature."
#         closer = "Please let us know if you have any questions — we're happy to help."
#     elif reminder_number <= 5:
#         subject = f"Following up — Fee Proposal {proposal_ref}"
#         opener = f"We're following up on the fee proposal we sent you for the above project. If you're still considering the engagement, we'd love to hear from you."
#         closer = "If the timing isn't right or you'd like to discuss the scope, please feel free to reach out."
#     else:
#         subject = f"Final follow-up — Fee Proposal {proposal_ref}"
#         opener = f"This is our final follow-up regarding the fee proposal for the above project. We'll close the matter from our end if we don't hear back."
#         closer = "If you'd like to revisit this engagement in the future, please don't hesitate to get in touch."

#     body = _base_html(f"""
#         <p>Hi {client_name},</p>
#         <p>{opener}</p>
#         <div class="detail">
#           <b>Proposal Ref:</b> {proposal_ref}<br>
#           <b>Project:</b>      {project_address}
#         </div>
#         <p><a class="btn" href="{signing_url}">Review &amp; Sign Proposal</a></p>
#         <p>Or copy this link: <a href="{signing_url}" style="color:#2e75b6;font-size:13px;">{signing_url}</a></p>
#         <p>{closer}</p>
#         <p>Kind regards,<br><b>{sender_name}</b><br>{sender_title}<br>{FIRM_NAME}</p>
#     """)

#     return _send(client_email, subject, body)

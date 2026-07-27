"""
signing.py — Fill an engineering proposal PDF's {{placeholder}} fields and
stamp the client's signature.

The template's text layer contains two categories of placeholder:

  Filled at PROPOSAL-GENERATION time, via docxtpl/Jinja (app.py's
  build_context() / smart_quote_sign()):
      {{client_name}}, {{project_type}}, {{project_address}}, {{fee}}
  These are known before the client ever sees the document, so they're
  real values by the time this module ever touches the PDF.

  Filled at ACTUAL SIGNING time, by this module:
      {{signature}}   — the drawn/typed signature
      {{date}}        — the date signed
  These are only known once the client signs, so the master template uses
  [[signature]] / [[date]] (inert to Jinja) which app.py's
  inject_deferred_signing_tags() converts to literal "{{signature}}" /
  "{{date}}" text AFTER docxtpl renders — so they survive PDF conversion
  untouched and can be found here.

This module finds every occurrence of "{{date}}" and "{{signature}}"
directly, redacts the raw placeholder text, and stamps the real value in
its exact place. This is layout-independent: it doesn't matter which page
a token lands on.
"""

import base64
from pathlib import Path
from datetime import datetime

import fitz  # PyMuPDF

DATE_FONT_SIZE = 11
TYPED_SIG_FONT_SIZE = 24  # generous size — the template reserves real vertical room now

# master_proposal_template.docx gives the "Signature: [[signature]]" paragraph
# 36pt of space_after (see check_template.py / the template's paragraph
# formatting), so there's genuine room below the placeholder's own line for a
# legible signature — not just the ~16pt of a single text line.
SIG_WIDTH = 200
SIG_TOP_OVERHANG = 2  # small rise above the placeholder's own top
SIG_BOTTOM_OVERHANG = 26  # extends down into the reserved space_after gap


def _decode_data_url(data_url: str) -> bytes:
    """Decode a 'data:image/png;base64,....' string into raw bytes."""
    if "," in data_url:
        _, encoded = data_url.split(",", 1)
    else:
        encoded = data_url
    return base64.b64decode(encoded)


def _find_all(doc: fitz.Document, token: str):
    """Return [(page_index, rect), ...] for every occurrence of token in the doc."""
    hits = []
    for page_index in range(len(doc)):
        for rect in doc[page_index].search_for(token):
            hits.append((page_index, rect))
    return hits


def _replace_text(
    doc,
    token,
    value,
    *,
    fontsize=DATE_FONT_SIZE,
    fontname="helv",
    color=(0.1, 0.1, 0.1),
):
    """
    Find every occurrence of `token`, redact it (blanks the underlying
    "{{...}}" text), and stamp `value` in its place. Returns the hits found,
    so callers can detect a missing placeholder.
    """
    hits = _find_all(doc, token)
    if not hits:
        print(f"[signing] WARNING: placeholder {token!r} not found — nothing replaced.")
        return hits
    for page_index, rect in hits:
        page = doc[page_index]
        page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()
        page.insert_text(
            fitz.Point(rect.x0, rect.y1),
            value,
            fontsize=fontsize,
            fontname=fontname,
            color=color,
        )
    return hits


def stamp_signature(
    source_pdf_path: str,
    output_pdf_path: str,
    *,
    sig_type: str,
    sig_data: str,
    signer_name: str,
    signer_position: str = "",
    signed_date: str = None,
    total_fee: str = "",
    project_type: str = "",
    project_address: str = "",
) -> None:
    """
    Stamps the client's signature and fills the "{{date}}" placeholder in
    the proposal PDF.

    sig_type: "draw" (sig_data is a data:image/png;base64,... string)
              or "type" (sig_data is the typed name as plain text)
    signer_name / signer_position / total_fee / project_type /
    project_address: kept as parameters for backward compatibility with
              existing callers, but are NOT stamped here anymore —
              client_name, project_type, project_address, and fee are now
              filled by docxtpl directly at proposal-generation time (see
              build_context() / smart_quote_sign() in app.py), so by the
              time a PDF reaches this function those fields already show
              real values, not literal "{{...}}" placeholders.
    """
    if signed_date is None:
        signed_date = datetime.now().strftime("%d %B %Y")

    doc = fitz.open(source_pdf_path)

    # 1. Fill the date placeholder (only remaining deferred text field)
    _replace_text(doc, "{{date}}", signed_date, fontsize=DATE_FONT_SIZE)

    # 2. Stamp the signature directly over the "{{signature}}" placeholder
    sig_hits = _find_all(doc, "{{signature}}")
    if not sig_hits:
        doc.close()
        raise RuntimeError(
            "Could not find '{{signature}}' placeholder in the PDF — cannot "
            "place signature. Check that master_proposal_template.docx "
            "uses '[[signature]]' (double square brackets) at this spot, "
            "and that inject_deferred_signing_tags() ran during generation."
        )
    sig_page_idx, sig_rect = sig_hits[0]
    page = doc[sig_page_idx]
    page.add_redact_annot(sig_rect, fill=(1, 1, 1))
    page.apply_redactions()

    img_rect = fitz.Rect(
        sig_rect.x0,
        sig_rect.y0 - SIG_TOP_OVERHANG,
        sig_rect.x0 + SIG_WIDTH,
        sig_rect.y1 + SIG_BOTTOM_OVERHANG,
    )

    if sig_type == "draw":
        img_bytes = _decode_data_url(sig_data)
        page.insert_image(img_rect, stream=img_bytes, keep_proportion=True)
    else:  # "type" — render the typed name in an italic script-style font
        # Anchor near the bottom of the reserved box, leaving room below the
        # baseline for descenders (g/y/j) and room above for ascenders at
        # this larger font size.
        baseline_y = img_rect.y1 - 6
        page.insert_text(
            fitz.Point(sig_rect.x0, baseline_y),
            sig_data,
            fontsize=TYPED_SIG_FONT_SIZE,
            # "heit" = Helvetica-Oblique. NOTE: "helv-oblique" is NOT a valid
            # PyMuPDF font name/alias and raises `Exception: need font file
            # or buffer` for every typed signature — must stay "heit".
            fontname="heit",
            color=(0.12, 0.18, 0.35),
        )

    Path(output_pdf_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_pdf_path)
    doc.close()
    print(f"[signing] PDF successfully filled and signed → {output_pdf_path}")

"""
Unit tests for signing.py, isolated from the rest of the app.

Builds a minimal PDF containing "{{signature}}" / "{{date}}" placeholder
text directly via PyMuPDF (not through the full docx pipeline) so these
tests are fast and don't depend on docx2pdf/LibreOffice being installed.
"""

import fitz
import pytest

import signing
from tests.conftest import VALID_SIG_DATA_URL


def _build_placeholder_pdf(path):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 200), "Signature: {{signature}}", fontsize=11)
    page.insert_text((72, 220), "Jane Client          Date: {{date}}", fontsize=11)
    doc.save(path)
    doc.close()


def test_stamp_draw_signature_succeeds(tmp_path):
    src = tmp_path / "source.pdf"
    out = tmp_path / "signed.pdf"
    _build_placeholder_pdf(str(src))

    signing.stamp_signature(
        str(src),
        str(out),
        sig_type="draw",
        sig_data=VALID_SIG_DATA_URL,
        signer_name="Jane Client",
        signer_position="",
        signed_date="16 July 2026",
    )

    assert out.exists()
    doc = fitz.open(str(out))
    text = doc[0].get_text()
    doc.close()
    # Placeholders must be gone — replaced with real content
    assert "{{signature}}" not in text
    assert "{{date}}" not in text
    assert "16 July 2026" in text


def test_stamp_typed_signature_succeeds(tmp_path):
    """
    Regression test: signing.py previously used fontname="helv-oblique",
    which is not a valid PyMuPDF font alias and raised
    `Exception: need font file or buffer` for every typed signature.
    The correct alias is "heit". This must never regress.
    """
    src = tmp_path / "source.pdf"
    out = tmp_path / "signed.pdf"
    _build_placeholder_pdf(str(src))

    signing.stamp_signature(
        str(src),
        str(out),
        sig_type="type",
        sig_data="Jane Client",
        signer_name="Jane Client",
        signer_position="",
        signed_date="16 July 2026",
    )

    assert out.exists()
    doc = fitz.open(str(out))
    text = doc[0].get_text()
    doc.close()
    assert "{{signature}}" not in text
    assert "Jane Client" in text


def test_stamp_signature_missing_placeholder_raises(tmp_path):
    src = tmp_path / "no_placeholder.pdf"
    out = tmp_path / "signed.pdf"

    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 200), "This document has no signature placeholder.", fontsize=11)
    doc.save(str(src))
    doc.close()

    with pytest.raises(RuntimeError, match="signature"):
        signing.stamp_signature(
            str(src),
            str(out),
            sig_type="draw",
            sig_data=VALID_SIG_DATA_URL,
            signer_name="Jane Client",
        )


def test_stamp_signature_invalid_image_data_raises(tmp_path):
    src = tmp_path / "source.pdf"
    out = tmp_path / "signed.pdf"
    _build_placeholder_pdf(str(src))

    with pytest.raises(Exception):
        signing.stamp_signature(
            str(src),
            str(out),
            sig_type="draw",
            sig_data="data:image/png;base64,not-valid-base64!!!",
            signer_name="Jane Client",
        )

"""
calibrate_signature_box.py — One-off diagnostic.

Draws a coordinate grid (every 50pt) onto the LAST page of a given PDF,
so you can visually read off the exact x/y position of the
"Acceptance of Proposal" signature table.

Usage:
    python calibrate_signature_box.py path\to\some_unsigned_proposal.pdf

Output:
    Creates "<input>_grid.pdf" next to the input file. Open it and find
    where the Signature / Name / Position / Date boxes sit on the grid.
    Report back the x,y values for each box's bottom-left corner.
"""

import sys
import io
from pathlib import Path

from reportlab.pdfgen import canvas as rl_canvas
from pypdf import PdfReader, PdfWriter


def build_grid_overlay(page_w: float, page_h: float) -> io.BytesIO:
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))

    c.setStrokeColorRGB(1, 0, 0)
    c.setFillColorRGB(1, 0, 0)
    c.setFont("Helvetica", 6)

    step = 50
    x = 0
    while x <= page_w:
        c.line(x, 0, x, page_h)
        c.drawString(x + 2, 4, str(x))
        x += step

    y = 0
    while y <= page_h:
        c.line(0, y, page_w, y)
        c.drawString(2, y + 2, str(y))
        y += step

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


def main():
    if len(sys.argv) < 2:
        print("Usage: python calibrate_signature_box.py <path-to-pdf>")
        sys.exit(1)

    src_path = Path(sys.argv[1])
    if not src_path.exists():
        print(f"File not found: {src_path}")
        sys.exit(1)

    reader = PdfReader(str(src_path))
    writer = PdfWriter()

    last_index = len(reader.pages) - 1

    for i, page in enumerate(reader.pages):
        if i == last_index:
            page_box = page.mediabox
            page_w = float(page_box.width)
            page_h = float(page_box.height)
            print(f"Last page index={i}, size={page_w} x {page_h}")

            overlay_buf = build_grid_overlay(page_w, page_h)
            overlay_reader = PdfReader(overlay_buf)
            page.merge_page(overlay_reader.pages[0])

        writer.add_page(page)

    out_path = src_path.with_name(src_path.stem + "_grid.pdf")
    with open(out_path, "wb") as f:
        writer.write(f)

    print(f"\n✅ Grid overlay saved to: {out_path}")
    print("Open it, find the Signature/Name/Position/Date table, and report")
    print("the approximate x,y (bottom-left corner) of each box.")


if __name__ == "__main__":
    main()

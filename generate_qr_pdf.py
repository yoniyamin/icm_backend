"""
=====================================================
QR Code PDF Generator with Cutting Guides
=====================================================

📌 Description:
This script generates a **printable PDF** containing QR codes in a **grid format**.
It minimizes padding, aligns QR codes perfectly for **cutting efficiency**, and
includes **light gray cutting guides**.

📌 Features:
✔ Reads QR codes from a specified directory.
✔ Supports **selecting a range** of QR codes by number.
✔ Aligns QR codes in a **4x5 grid per page** (modifiable).
✔ **No padding between rows** for efficient cutting.
✔ **Adds column separation for cutting precision**.
✔ Saves the output as a **printable PDF**.

📌 File Naming Convention:
- The script expects QR code filenames in the format: `qr_for_book_<number>.png`
  Example: `qr_for_book_1.png`, `qr_for_book_2.png`, `qr_for_book_10.png`

📌 Dependencies:
Make sure you have the required libraries installed:
- `fpdf`: For PDF generation.

📌 Usage:
- `python generate_qr_pdf.py <qr_directory> <start_qr> <end_qr>`
- `qr_directory` is the directory containing the QR code images.
- `start_qr` and `end_qr` are the range of QR codes to generate.

📌 Example:
- `python generate_qr_pdf.py qr_codes 1 10` will generate a PDF with QR codes 1 to 10.
"""

import os
import re
import sys
from fpdf import FPDF

# PDF layout settings
QR_SIZE = 50  # QR code image size in mm
COLUMN_PADDING = 0  # Small gap between columns for cutting
COLUMNS = 4  # Number of QR codes per row
ROWS = 5  # Number of QR codes per column per page
PAGE_WIDTH = 210  # A4 width in mm
PAGE_HEIGHT = 297  # A4 height in mm
START_X = (PAGE_WIDTH - (COLUMNS * QR_SIZE + (COLUMNS - 1) * COLUMN_PADDING)) / 2
START_Y = 20  # Top margin

def extract_qr_number(filename):
    """Extracts the numeric part of the QR filename (e.g., qr_for_book_1.png -> 1)."""
    match = re.search(r"qr_for_book_(\d+)\.png", filename)
    return int(match.group(1)) if match else None

def generate_qr_pdf(qr_directory, output_pdf, start_qr, end_qr):
    """Generates a printable PDF with QR codes arranged in a grid layout with cutting guides."""
    # Get list of QR code images
    qr_files = sorted([
        os.path.join(qr_directory, f) for f in os.listdir(qr_directory)
        if f.endswith(".png") and extract_qr_number(f) is not None
    ], key=lambda x: extract_qr_number(os.path.basename(x)))

    # Filter by range
    qr_files = [f for f in qr_files if start_qr <= extract_qr_number(os.path.basename(f)) <= end_qr]

    if not qr_files:
        print(f"No QR codes found in the range {start_qr}-{end_qr}.")
        return

    # Create PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    # Add first page
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    for i, qr_path in enumerate(qr_files):
        row = (i // COLUMNS) % ROWS
        col = i % COLUMNS

        x = START_X + col * (QR_SIZE + COLUMN_PADDING)
        y = START_Y + row * QR_SIZE  # No padding between rows

        # Add QR code image
        pdf.image(qr_path, x=x, y=y, w=QR_SIZE, h=QR_SIZE)

        # Add cutting guides
        pdf.set_draw_color(200, 200, 200)  # Light gray cutting lines
        pdf.line(x, y, x + QR_SIZE, y)  # Top line
        pdf.line(x, y + QR_SIZE, x + QR_SIZE, y + QR_SIZE)  # Bottom line
        pdf.line(x, y, x, y + QR_SIZE)  # Left line
        pdf.line(x + QR_SIZE, y, x + QR_SIZE, y + QR_SIZE)  # Right line

        # Add a new page when the grid is filled
        if (i + 1) % (COLUMNS * ROWS) == 0 and i + 1 < len(qr_files):
            pdf.add_page()

    # Save PDF
    pdf.output(output_pdf)
    print(f"✅ QR codes PDF generated: {output_pdf}")

# Run the script
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python generate_qr_pdf.py <qr_directory> <start_qr> <end_qr>")
    else:
        qr_directory = sys.argv[1]
        start_qr = int(sys.argv[2])
        end_qr = int(sys.argv[3])
        output_pdf = "qr_codes_printable.pdf"

        generate_qr_pdf(qr_directory, output_pdf, start_qr, end_qr)

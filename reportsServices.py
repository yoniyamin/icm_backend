import os
import re
from fpdf import FPDF

# Fixed QR codes directory for the application.
QR_CODES_DIR = os.path.join(os.getcwd(), "qr_codes")

# PDF layout settings
QR_SIZE = 50            # QR code image size in mm
COLUMN_PADDING = 0      # Gap between columns (for cutting guides)
COLUMNS = 4             # Number of QR codes per row
ROWS = 5                # Number of QR codes per page (rows)
PAGE_WIDTH = 210        # A4 width in mm
PAGE_HEIGHT = 297       # A4 height in mm
START_X = (PAGE_WIDTH - (COLUMNS * QR_SIZE + (COLUMNS - 1) * COLUMN_PADDING)) / 2
START_Y = 20            # Top margin

def extract_qr_number(filename):
    """
    Extracts the numeric part from a QR code filename.
    For example: "qr_for_book_1.png" -> 1
    """
    match = re.search(r"qr_for_book_(\d+)\.png", filename)
    return int(match.group(1)) if match else None

def generate_qr_pdf(start_qr, end_qr):
    """
    Generates a printable PDF with QR codes in a grid (with cutting guides)
    and returns the PDF as bytes.

    Parameters:
      - start_qr (int): The starting QR code number.
      - end_qr (int): The ending QR code number.

    Returns:
      - bytes: The PDF content.

    Raises:
      - ValueError: If no QR codes are found in the specified range.
    """
    # Use the fixed directory for QR codes.
    qr_directory = QR_CODES_DIR

    # Gather and sort QR code image file paths from the fixed directory.
    qr_files = sorted(
        [os.path.join(qr_directory, f) for f in os.listdir(qr_directory)
         if f.endswith(".png") and extract_qr_number(f) is not None],
        key=lambda x: extract_qr_number(os.path.basename(x))
    )

    # Filter files by the given range.
    qr_files = [
        f for f in qr_files
        if start_qr <= extract_qr_number(os.path.basename(f)) <= end_qr
    ]

    if not qr_files:
        raise ValueError(f"No QR codes found in the range {start_qr}-{end_qr}.")

    # Create the PDF.
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    for i, qr_path in enumerate(qr_files):
        row = (i // COLUMNS) % ROWS
        col = i % COLUMNS

        x = START_X + col * (QR_SIZE + COLUMN_PADDING)
        y = START_Y + row * QR_SIZE  # No vertical padding

        # Add the QR code image.
        pdf.image(qr_path, x=x, y=y, w=QR_SIZE, h=QR_SIZE)

        # Draw cutting guides (light gray lines).
        pdf.set_draw_color(200, 200, 200)
        pdf.line(x, y, x + QR_SIZE, y)                   # Top line
        pdf.line(x, y + QR_SIZE, x + QR_SIZE, y + QR_SIZE)  # Bottom line
        pdf.line(x, y, x, y + QR_SIZE)                   # Left line
        pdf.line(x + QR_SIZE, y, x + QR_SIZE, y + QR_SIZE)  # Right line

        # Add a new page when the grid is filled.
        if (i + 1) % (COLUMNS * ROWS) == 0 and (i + 1) < len(qr_files):
            pdf.add_page()

    # Return the PDF as bytes.
    pdf_bytes = pdf.output(dest="S").encode("latin1")
    return pdf_bytes

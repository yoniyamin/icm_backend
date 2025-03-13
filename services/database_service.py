import io

import pandas as pd  # Add pandas import at the top
import sqlite3
import psycopg2
from psycopg2.extras import RealDictCursor
import os
from io import BytesIO
import qrcode
from datetime import date, datetime, timedelta, timezone
from PIL import Image, ImageDraw, ImageFont
from bidi.algorithm import get_display  # Correctly display RTL text
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from reportlab.lib.utils import ImageReader

QR_CODE_DIR = "qr_codes"

POSTGRES_URL = os.getenv("DATABASE_URL")
SQLITE_DB_PATH = "database.db"

def get_postgres_connection():
    """Establish a PostgreSQL connection."""
    return psycopg2.connect(POSTGRES_URL, cursor_factory=RealDictCursor)

def get_sqlite_connection():
    """Establish an SQLite connection (for sessions only)."""
    return sqlite3.connect(SQLITE_DB_PATH)

def check_neon_db_health():
    """Check if NeonDB is reachable by running a simple query."""
    try:
        with get_postgres_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT 1;")
                cursor.fetchone()
        return True  # Database is reachable
    except Exception as e:
        print(f"NeonDB health check failed: {e}")
        return False  # Database is unreachable

# ✅ SESSION MANAGEMENT (SQLite)
def store_session_token(token, expiry):
    print(f"DEBUG: Storing token {token} with expiry {expiry}")
    with get_sqlite_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
        INSERT INTO sessions (token, expiry) 
        VALUES (?, ?)
        ''', (token, expiry))
        conn.commit()
        return f"Token {token} stored with expiry {expiry}"
def validate_session_token(token):
    with get_sqlite_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
        SELECT expiry FROM sessions WHERE token = ?
        ''', (token,))
        result = cursor.fetchone()
        if result:
            expiry = datetime.fromisoformat(result[0])  # Token expiry in UTC
            current_time = datetime.now(timezone.utc)  # Current time in UTC
            print(f"DEBUG: Token expiry: {expiry}, Current time: {current_time}")
            if expiry > current_time:
                return True  # Token is valid
            else:
                print("DEBUG: Token has expired.")
    return False

def remove_expired_tokens():
    with get_sqlite_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM sessions WHERE expiry < ?', (datetime.now(timezone.utc).isoformat(),))
        deleted_rows = cursor.rowcount  # Get the number of rows deleted
        conn.commit()
        print(f"Expired tokens cleaned up: {deleted_rows}")
        return deleted_rows

# ✅ BOOK MANAGEMENT (PostgreSQL)
def get_books(order_by="desc"):
    """Retrieve books with correct loan status and borrower details."""
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cursor:
            order_clause = "DESC" if order_by.lower() == "desc" else "ASC"

            cursor.execute(f"""
                SELECT 
                    books.*,
                    COALESCE(loans.borrowed_at, NULL) AS borrowed_at,
                    COALESCE(members.parent_name, NULL) AS borrowing_child,
                    CASE 
                        WHEN loans.id IS NOT NULL THEN 'borrowed'
                        ELSE 'available'
                    END AS loan_status
                FROM books
                LEFT JOIN loans ON books.id = loans.book_id AND loans.returned_at IS NULL
                LEFT JOIN members ON loans.member_id = members.id
                ORDER BY books.created_at {order_clause}
            """)

            books = cursor.fetchall()
            print("📚 Books fetched:", books)  # Debugging
            return books

def convert_to_int_or_none(val):
    """
    Convert an incoming value to int if possible.
    If it's None, empty string, or invalid, return None instead.
    """
    if val is None:
        return None
    if isinstance(val, int):
        return val
    val_str = str(val).strip()
    if not val_str:
        # empty string => None
        return None
    try:
        return int(val_str)
    except ValueError:
        # can't parse => None or raise an error
        return None

def add_book(title, author, description, year_of_publication, cover_type, pages,
             recommended_age, book_condition, loan_status, delivering_parent):
    """
    Insert a new book into PostgreSQL and store its generated QR code in the qr_codes table.
    """
    qr_code_path = None
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            year_of_publication = convert_to_int_or_none(year_of_publication)
            pages = convert_to_int_or_none(pages)
            recommended_age = convert_to_int_or_none(recommended_age)

            try:
                # Insert the book with a temporary QR code.
                cursor.execute('''
                    INSERT INTO books (qr_code, title, author, description, year_of_publication, cover_type, pages,
                    recommended_age, book_condition, loan_status, delivering_parent)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                ''', (
                    "temp_qr_code", title, author, description, year_of_publication, cover_type, pages,
                    recommended_age, book_condition, loan_status, delivering_parent))
                book_id = cursor.fetchone()['id']
                # Generate the actual QR code string and image.
                qr_code = f"qr_for_book_{book_id}"
                image_data = generate_qr_code_with_logo(qr_code, title)
                # Insert the QR code image into the qr_codes table.
                cursor.execute('''
                    INSERT INTO qr_codes (qr_code, image)
                    VALUES (%s, %s)
                ''', (qr_code, psycopg2.Binary(image_data)))
                # Update the book record with the final QR code.
                cursor.execute("UPDATE books SET qr_code = %s WHERE id = %s", (qr_code, book_id))
                conn.commit()
                return qr_code
            except Exception as e:
                conn.rollback()
                raise e

def store_qr_code_in_db(qr_code, image_bytes):
    """Store the generated QR code image bytes in the database."""
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO qr_codes (qr_code, image)
                VALUES (%s, %s)
                ON CONFLICT (qr_code) DO UPDATE SET image = EXCLUDED.image
            """, (qr_code, psycopg2.Binary(image_bytes)))
            conn.commit()

def download_qr_code(qr_code):
    """
    Retrieve the QR code image binary data from the database given a qr_code identifier.
    """
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute("SELECT image FROM qr_codes WHERE qr_code = %s", (qr_code,))
            result = cursor.fetchone()
            if result and result.get("image"):
                return result["image"]
            else:
                return None


def generate_qr_code_with_logo(qr_code, title):
    """
    Generate a QR code image with a logo and title. Returns the PNG image bytes.
    """
    # Convert the title for proper RTL display if needed.
    rtl_title = get_display(title)

    # Create the QR code.
    qr = qrcode.QRCode(
        version=None,  # Automatically choose version
        error_correction=qrcode.constants.ERROR_CORRECT_Q,
        box_size=12,
        border=5,
    )
    qr.add_data(qr_code)
    qr.make(fit=True)
    print(f"QR Code Version: {qr.version}")

    # Create the QR image.
    qr_img = qr.make_image(fill="black", back_color="white").convert("RGB")
    qr_base_size = qr_img.size[0]

    # Try to load and overlay the logo.
    here = os.path.abspath(os.path.dirname(__file__))
    logo_path = os.path.join(here, "..", "static", "icm_logo.png")
    try:
        logo = Image.open(logo_path)
        logo_size = int(qr_base_size * 0.20)
        logo = logo.resize((logo_size, logo_size), Image.LANCZOS)
        logo_bg = Image.new('RGBA', (logo_size + 8, logo_size + 8), 'white')
        logo_pos = ((logo_bg.size[0] - logo_size) // 2, (logo_bg.size[1] - logo_size) // 2)
        logo_bg.paste(logo, logo_pos, mask=logo if logo.mode == 'RGBA' else None)
        pos = ((qr_base_size - logo_bg.size[0]) // 2, (qr_base_size - logo_bg.size[1]) // 2)
        mask = Image.new('L', qr_img.size, 255)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rectangle([pos[0], pos[1], pos[0] + logo_bg.size[0], pos[1] + logo_bg.size[1]], fill=0)
        qr_img.paste(logo_bg, pos, mask=logo_bg if logo_bg.mode == 'RGBA' else None)
    except Exception as e:
        print(f"Error adding logo to QR code: {e}")

    # Add space at the bottom for the title.
    title_space = 50
    canvas = Image.new("RGB", (qr_img.size[0], qr_img.size[1] + title_space), "white")
    canvas.paste(qr_img, (0, 0))
    draw = ImageDraw.Draw(canvas)
    font_path = os.path.join(here,"..", "static", "FreeSans.ttf")
    try:
        try:
            font = ImageFont.truetype(font_path, 24)
            print("Font file found:", font_path)
            for char in title:
                if font.getmask(char).getbbox() is None:
                    print(f"Warning: Character '{char}' not supported by font")
        except IOError:
            print("Font file not found:", font_path)
            font = ImageFont.load_default()
        except Exception as e:
            print(f"Font error: {e}")
            font = ImageFont.load_default()
        text_bbox = draw.textbbox((0, 0), rtl_title, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        text_position = ((canvas.size[0] - text_width) // 2, qr_img.size[1] + (title_space - text_height) // 2)
        # Draw a simple outline for better readability.
        draw.text((text_position[0] - 1, text_position[1]), rtl_title, fill="white", font=font)
        draw.text((text_position[0] + 1, text_position[1]), rtl_title, fill="white", font=font)
        draw.text((text_position[0], text_position[1] - 1), rtl_title, fill="white", font=font)
        draw.text((text_position[0], text_position[1] + 1), rtl_title, fill="white", font=font)
        draw.text(text_position, title, fill="black", font=font, direction='rtl')
    except Exception as e:
        print(f"Error adding title: {e}")

    # Save the image to a BytesIO buffer.
    img_buffer = io.BytesIO()
    canvas.save(img_buffer, format="PNG", quality=95)
    img_buffer.seek(0)
    image_data = img_buffer.read()
    print(f"Generated QR code for {qr_code}, image size in bytes: {len(image_data)}")
    return image_data

def get_all_qr_codes_with_title():
    """
    Return a list of dicts: [
      { 'qr_code': ..., 'title': ..., 'image': ... },
      ...
    ]
    including the book's title if it exists.
    """
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            # We LEFT JOIN books on matching qr_code
            # so that if a qr_code in 'qr_codes' doesn't have an entry in 'books', it still shows up.
            cursor.execute("""
                SELECT qr_codes.qr_code,
                       books.title
                FROM qr_codes
                LEFT JOIN books ON qr_codes.qr_code = books.qr_code
                ORDER BY qr_codes.qr_code
            """)
            return cursor.fetchall()


def generate_qr_pdf_report_by_list(qr_code_list):
    """
    Generate a PDF containing QR codes arranged in a 3-column, 4-row grid
    with cutting guides, each QR code bigger than before.
    """
    import io
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    from PIL import Image
    from reportlab.lib.utils import ImageReader

    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            query = """
                SELECT qr_codes.qr_code, qr_codes.image
                FROM qr_codes
                WHERE qr_codes.qr_code = ANY(%s)
                ORDER BY qr_codes.qr_code
            """
            cursor.execute(query, (qr_code_list,))
            qr_codes = cursor.fetchall()

    if not qr_codes:
        raise Exception("No matching QR codes found.")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # Use 3 columns instead of 4 for bigger images
    cols = 4
    # Keep 4 rows
    rows = 4
    # Margin around the page
    margin = 0.5 * inch

    # Calculate the available cell space
    qr_width = (width - 2 * margin) / cols
    qr_height = (height - 2 * margin) / rows

    # The actual QR code is square
    qr_size = min(qr_width, qr_height)

    # Center the grid horizontally
    x_start = margin + (width - 2 * margin - (cols * qr_size)) / 2
    y_start = height - margin

    qr_index = 0
    total_qrs = len(qr_codes)

    # We'll reduce the padding from 10% down to 5%
    padding_ratio = 0.05  # 5%

    while qr_index < total_qrs:
        # Draw dashed cutting guides
        c.setStrokeColorRGB(0.8, 0.8, 0.8)  # Light gray
        c.setDash([2, 2])

        # Horizontal lines
        for row in range(rows + 1):
            y_line = height - margin - (row * qr_size)
            c.line(margin, y_line, width - margin, y_line)

        # Vertical lines
        for col in range(cols + 1):
            x_line = x_start + (col * qr_size)
            c.line(x_line, margin, x_line, height - margin)

        # Draw each QR code
        for row in range(rows):
            for col in range(cols):
                if qr_index >= total_qrs:
                    break

                qr = qr_codes[qr_index]

                # Position for this cell
                x = x_start + (col * qr_size)
                y = y_start - (row * qr_size)

                # Convert binary to a BytesIO for PIL
                img_buffer = io.BytesIO(qr["image"])

                # Use 5% padding inside each cell
                pad = qr_size * padding_ratio
                c.drawImage(
                    ImageReader(img_buffer),
                    x + pad,
                    y - qr_size + pad,
                    width=qr_size - 2 * pad,
                    height=qr_size - 2 * pad
                )

                qr_index += 1

            if qr_index >= total_qrs:
                break

        # If there are more QRs left, move to a new page
        if qr_index < total_qrs:
            c.showPage()
            y_start = height - margin

    c.save()
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def update_book_status(qr_code, status):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("UPDATE books SET status = %s WHERE qr_code = %s", (status, qr_code))
            conn.commit()

def get_book_loans(book_id):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute("SELECT * FROM loans WHERE book_id = %s", (book_id,))
            loans = cursor.fetchall()
            return [dict(loan) for loan in loans]

def borrow_book(qr_code, member_id, borrowed_date, book_state):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            # Lookup book_id from qr_code
            cursor.execute("SELECT id FROM books WHERE qr_code = %s", (qr_code,))
            book = cursor.fetchone()

            if not book:
                return False

            book_id = book['id']  # Use dictionary key 'id' instead of index 0

            # Insert a new loan record
            cursor.execute("""
                INSERT INTO loans (book_id, member_id, borrowed_at, book_state)
                VALUES (%s, %s, %s, %s)
            """, (book_id, member_id, borrowed_date, book_state))

            # Update the loan_status in the books table to 'borrowed'
            cursor.execute("""
                UPDATE books SET loan_status = 'borrowed' WHERE id = %s
            """, (book_id,))

            try:
                conn.commit()
            except Exception as e:
                conn.rollback()
                return False

    return True


def update_book(book_id, **kwargs):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            try:
                # Filter allowed fields and prepare update parameters
                allowed_fields = {
                    'title', 'author', 'description', 'year_of_publication',
                    'cover_type', 'pages', 'recommended_age', 'book_condition',
                    'delivering_parent'
                }


                update_fields = []
                values = []

                for field, value in kwargs.items():
                    if field in allowed_fields:
                        if field in ("year_of_publication", "pages", "recommended_age"):
                            value = convert_to_int_or_none(value)

                        update_fields.append(f"{field} = %s")
                        values.append(value)

                if not update_fields:
                    return None  # No valid fields to update

                # Add book_id as the last parameter
                values.append(book_id)

                # Build the update query
                query = f'''
                    UPDATE books 
                    SET {', '.join(update_fields)}, updated_at = CURRENT_TIMESTAMP
                    WHERE id = %s
                    RETURNING id, title, author, description, year_of_publication, cover_type, 
                            pages, recommended_age, book_condition, loan_status, delivering_parent, qr_code
                '''

                cursor.execute(query, values)
                conn.commit()

                # Return the updated book
                updated_book = cursor.fetchone()
                if updated_book:
                    return dict(updated_book)
                return None

            except Exception as e:
                print(f"Database error: {str(e)}")
                conn.rollback()
                return None


def return_book(qr_code):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT id FROM books WHERE qr_code = %s", (qr_code,))
            book = cursor.fetchone()
            if not book:
                return {"success": False, "message": "Book not found"}
            book_id = book['id']
            try:
                cursor.execute("""
                    UPDATE loans 
                    SET returned_at = CURRENT_TIMESTAMP 
                    WHERE book_id = %s AND returned_at IS NULL
                """, (book_id,))
                cursor.execute("""
                    SELECT COUNT(*) FROM loans WHERE book_id = %s AND returned_at IS NULL
                """, (book_id,))
                active_loans_count = cursor.fetchone()['count']
                if active_loans_count == 0:
                    cursor.execute("""
                        UPDATE books SET loan_status = 'available' WHERE id = %s
                    """, (book_id,))
                conn.commit()
            except Exception as e:
                conn.rollback()
                # Fallback: Sync status to prevent inconsistency
                cursor.execute("""
                    UPDATE books 
                    SET loan_status = 'available'
                    WHERE id = %s AND id NOT IN (
                        SELECT book_id FROM loans WHERE returned_at IS NULL
                    )
                """, (book_id,))
                conn.commit()
                return {"success": False, "message": f"Error returning book, status corrected: {str(e)}"}
    return {"success": True, "message": "Book returned successfully"}


def get_members():
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cursor:
            cursor.execute("""
                SELECT m.*, COALESCE(COUNT(l.id), 0) AS borrowed_books_count
                FROM members m
                LEFT JOIN loans l 
                    ON m.id = l.member_id AND l.returned_at IS NULL
                GROUP BY m.id
                ORDER BY borrowed_books_count DESC, m.created_at DESC
            """)
            return cursor.fetchall()


def add_member(parent_name, kid_name, email):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute('''
            INSERT INTO members (parent_name, kid_name, email)
            VALUES (%s, %s, %s)
            ''', (parent_name, kid_name, email))
            conn.commit()

def update_member(member_id, parent_name, kid_name, email):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute('''
                UPDATE members 
                SET parent_name = %s, kid_name = %s, email = %s, updated_at = CURRENT_TIMESTAMP
                WHERE id = %s
            ''', (parent_name, kid_name, email, member_id))
            conn.commit()

def delete_member(member_id):
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            # Check if there are any open loans for this member
            cursor.execute(
                "SELECT COUNT(*) FROM loans WHERE member_id = %s AND returned_at IS NULL",
                (member_id,)
            )
            open_loans_count = cursor.fetchone()['count']  # RealDictCursor returns dict

            if open_loans_count > 0:
                raise Exception("Cannot delete member with open loans.")

            # Proceed to delete the member if no open loans
            cursor.execute('DELETE FROM members WHERE id = %s', (member_id,))
            conn.commit()

def get_book_by_qr_code(qr_code):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute("SELECT * FROM books WHERE qr_code = %s", (qr_code,))
            book = cursor.fetchone()
            return dict(book) if book else None


def get_books_by_status(param):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            if param == 'borrowed':
                cursor.execute("""
                    SELECT 
                        books.*, 
                        loans.borrowed_at, 
                        members.parent_name AS borrowing_child
                    FROM books
                    JOIN loans ON books.id = loans.book_id AND loans.returned_at IS NULL
                    LEFT JOIN members ON loans.member_id = members.id
                """)
            else:  # 'available'
                cursor.execute("""
                    SELECT books.*
                    FROM books
                    WHERE books.id NOT IN (
                        SELECT book_id FROM loans WHERE returned_at IS NULL
                    )
                """)
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

def get_borrowing_history(qr_code=None):
    query = '''
    SELECT 
        books.title AS book_title,
        loans.book_id AS book_id,
        members.kid_name AS borrowed_name,
        loans.borrowed_at AS loan_start,
        loans.returned_at AS return_date,
        loans.book_state AS state
    FROM loans
    JOIN books ON books.id = loans.book_id
    JOIN members ON members.id = loans.member_id
    '''
    params = []
    if qr_code:
        query += " WHERE books.qr_code = %s"
        params.append(qr_code)

    query += " ORDER BY loans.borrowed_at DESC"

    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute(query, params)
            results = cursor.fetchall()
            return [dict(row) for row in results]


def get_open_loans(qr_code=None):
    query = """
        SELECT 
            l.id AS loan_id,
            l.book_id,
            b.title AS book_title,
            m.parent_name AS borrower_name,
            l.borrowed_at AS loan_start_date,
            l.returned_at AS return_date
        FROM loans l
        JOIN books b ON l.book_id = b.id
        JOIN members m ON l.member_id = m.id
        WHERE l.returned_at IS NULL
    """
    params = []
    if qr_code:
        query += " AND b.qr_code = %s"
        params.append(qr_code)

    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]


def get_loan_history(qr_code, show_all):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            query = """
                SELECT l.id, l.book_id, l.borrowed_at, l.returned_at, l.book_state, b.title AS book_title, m.parent_name AS borrower_name, m.kid_name AS borrower_child
                FROM loans l
                JOIN books b ON l.book_id = b.id
                JOIN members m ON l.member_id = m.id
                WHERE b.qr_code = %s
            """
            if not show_all:
                query += " AND l.returned_at IS NULL"  # Filter for open loans only
            query += " ORDER BY l.borrowed_at DESC"

            cursor.execute(query, (qr_code,))
            loans = cursor.fetchall()
            return [dict(loan) for loan in loans]


def get_all_open_loans():
    print("getting only open loans")
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            query = """
                SELECT l.id, l.book_id, l.borrowed_at, l.returned_at, l.book_state, b.title AS book_title, m.parent_name AS borrower_name, m.kid_name AS borrower_child
                FROM loans l
                JOIN books b ON l.book_id = b.id
                JOIN members m ON l.member_id = m.id
                WHERE l.returned_at IS NULL  -- Only open loans
                ORDER BY l.borrowed_at DESC
            """
            cursor.execute(query)
            loans = cursor.fetchall()
            return [dict(loan) for loan in loans]

def get_all_loans():
    print("getting all loans")
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            query = """
                SELECT l.id, l.book_id, l.borrowed_at, l.returned_at, l.book_state, b.title AS book_title, m.parent_name AS borrower_name, m.kid_name AS borrower_child
                FROM loans l
                JOIN books b ON l.book_id = b.id
                JOIN members m ON l.member_id = m.id
                ORDER BY l.borrowed_at DESC
            """
            cursor.execute(query)
            loans = cursor.fetchall()
            return [dict(loan) for loan in loans]


# Function to extract book data for reporting
def get_books_report(order_by="desc", sort_column="title", include_history=True):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            # Validate the order_by parameter
            order_clause = "DESC" if order_by.lower() == "desc" else "ASC"
            valid_sort_columns = ["created_at", "borrowed_at", "title"]
            sort_column = sort_column if sort_column in valid_sort_columns else "title"

            # Fetch loan data based on include_history parameter
            if include_history:
                # Fetch all loans (both open and closed)
                query = f"""
                    SELECT 
                        books.title,
                        books.author,
                        loans.borrowed_at,
                        loans.returned_at,
                        members.parent_name AS borrowed_by,
                        members.email AS borrower_email
                    FROM books
                    LEFT JOIN loans ON books.id = loans.book_id
                    LEFT JOIN members ON loans.member_id = members.id
                    ORDER BY books.{sort_column} {order_clause}
                """
            else:
                # Fetch only open loans
                query = f"""
                    SELECT 
                        books.title,
                        books.author,
                        loans.borrowed_at,
                        members.parent_name AS borrowed_by,
                        members.email AS borrower_email
                    FROM books
                    LEFT JOIN loans ON books.id = loans.book_id AND loans.returned_at IS NULL
                    LEFT JOIN members ON loans.member_id = members.id
                    WHERE loans.returned_at IS NULL
                    ORDER BY books.{sort_column} {order_clause}
                """

            cursor.execute(query)
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

# Function to generate an Excel report
def generate_books_report(order_by="desc", sort_column="title", include_history=True, language="he"):
    # Extract the data using the `get_books_report` function
    books_data = get_books_report(order_by, sort_column, include_history)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(books_data)

    # **Dynamic Header Translation**
    header_translations = {
        'en': {
            'title': 'Title',
            'author': 'Author',
            'borrowed_at': 'Borrowed At',
            'returned_at': 'Returned At',
            'borrowed_by': 'Borrowed By',
            'borrower_email': 'Borrower Email'
        },
        'he': {
            'title': 'שם הספר',
            'author': 'שם המחבר',
            'borrowed_at': 'תאריך השאלה',
            'returned_at': 'תאריך החזרה',
            'borrowed_by': 'הושאל על ידי',
            'borrower_email': 'אימייל השואל'
        }
    }

    # Get headers based on the language
    headers = header_translations.get(language, header_translations['en'])
    translated_columns = [headers.get(col, col) for col in df.columns]
    df.columns = translated_columns  # Rename DataFrame columns to translated headers

    if language == 'he':
        df.columns = [get_display(col) for col in df.columns]  # Apply RTL for Hebrew

    # Generate the Excel file
    report_type = "Active Loans" if not include_history else "Historical Loans"
    report_filename = f"books_report_{report_type}_{date.today()}.xlsx"

    # Create a workbook and add the data
    wb = Workbook()
    ws = wb.active
    ws.title = "ספריית הקהילה הישראלית במדריד - " if language == 'he' else "ICM Library - "

    # Set header
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="4ECAC7", end_color="4ECAC7", fill_type="solid")
    report_title = "דוח השאלות" if language == 'he' else "Loans Report"
    ws.append([report_title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    header_cell = ws.cell(row=1, column=1)
    header_cell.font = Font(bold=True, size=14)
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Add data rows with borders
    border = Border(left=Side(style="thin", color="000000"),
                    right=Side(style="thin", color="000000"),
                    top=Side(style="thin", color="000000"),
                    bottom=Side(style="thin", color="000000"))

    # Add headers
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 2):
        if r_idx == 2:
            for c_idx, cell_value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
                cell.border = border
        else:
            for c_idx, cell_value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.border = border
                if include_history and df.at[r_idx - 3, 'returned_at'] is None:
                    cell.fill = PatternFill(start_color="FEC43C", end_color="FEC43C", fill_type="solid")

    # Freeze the header row
    ws.freeze_panes = "A3"

    # Add filters to columns (start from the second row)
    ws.auto_filter.ref = f"A2:{ws.cell(row=2, column=len(df.columns)).coordinate}"

    # Save the workbook
    wb.save(report_filename)

    return report_filename

# Function to generate an inventory report
def generate_inventory_report(order_by="desc", sort_column="title", include_borrowed=True, language="he"):
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            print("include_borrowed:", include_borrowed)

            # Validate the order_by parameter
            order_clause = "DESC" if order_by.lower() == "desc" else "ASC"
            valid_sort_columns = ["created_at", "title"]
            sort_column = sort_column if sort_column in valid_sort_columns else "title"

            # Fetch inventory data based on include_borrowed parameter
            if include_borrowed:
                query = f"""
                    SELECT 
                        books.id,
                        books.title,
                        books.author,
                        books.description,
                        books.year_of_publication,
                        books.pages,
                        books.cover_type,
                        books.book_condition,
                        books.loan_status
                    FROM books
                    ORDER BY books.{sort_column} {order_clause}
                """
            else:
                query = f"""
                    SELECT 
                        books.id,
                        books.title,
                        books.author,
                        books.description,
                        books.year_of_publication,
                        books.pages,
                        books.cover_type,
                        books.book_condition,
                        books.loan_status
                    FROM books
                    WHERE books.loan_status = 'available'
                    ORDER BY books.{sort_column} {order_clause}
                """

            cursor.execute(query)
            rows = cursor.fetchall()
            books_data = [dict(row) for row in rows]

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(books_data)

    # **Dynamic Header Translation**
    header_translations = {
        'en': {
            'id': 'ID',
            'title': 'Title',
            'author': 'Author',
            'description': 'Description',
            'year_of_publication': 'Year of Publication',
            'pages': 'Pages',
            'cover_type': 'Cover Type',
            'book_condition': 'Book Condition',
            'loan_status': 'Loan Status'
        },
        'he': {
            'id': 'מזהה',
            'title': 'שם הספר',
            'author': 'שם המחבר',
            'description': 'תיאור',
            'year_of_publication': 'שנת פרסום',
            'pages': 'עמודים',
            'cover_type': 'סוג הכריכה',
            'book_condition': 'מצב הספר',
            'loan_status': 'סטטוס השאלה'
        }
    }

    # Get headers based on the language
    headers = header_translations.get(language, header_translations['en'])
    translated_columns = [headers.get(col, col) for col in df.columns]
    df.columns = translated_columns  # Rename DataFrame columns to translated headers

    if language == 'he':
        df.columns = [get_display(col) for col in df.columns]  # Apply RTL for Hebrew

    # Generate the Excel file
    report_filename = f"inventory_report_{date.today()}.xlsx"

    # Create a workbook and add the data
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # Add report title
    report_title = "דוח מלאי" if language == 'he' else "Inventory Report"
    ws.append([report_title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    header_cell = ws.cell(row=1, column=1)
    header_cell.font = Font(bold=True, size=14)
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Add data rows with borders
    border = Border(left=Side(style="thin", color="000000"),
                    right=Side(style="thin", color="000000"),
                    top=Side(style="thin", color="000000"),
                    bottom=Side(style="thin", color="000000"))

    # Set header
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="4ECAC7", end_color="4ECAC7", fill_type="solid")
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 2):
        if r_idx == 2:  # Headers row
            for c_idx, cell_value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
                cell.border = border
        else:  # Data rows
            for c_idx, cell_value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.border = border

    # Freeze the header row
    ws.freeze_panes = "A3"

    # Add filters to columns (start from the second row)
    ws.auto_filter.ref = f"A2:{ws.cell(row=2, column=len(df.columns)).coordinate}"

    # Save the workbook
    wb.save(report_filename)

    return report_filename


def find_email_by_borrower_name(borrower_name):
    """
    Look up the email of a member by the borrower name.
    The borrower name corresponds to the `parent_name` in the members table.
    """
    try:
        with get_postgres_connection() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT email FROM members WHERE parent_name = %s", (borrower_name,))
                result = cursor.fetchone()
                if result:
                    email = result["email"]  # Extract the email from the result
                    return email
                else:
                    print(f"No email found for borrower name: {borrower_name}")
                    return None
    except Exception as e:
        print(f"Error finding email for borrower name '{borrower_name}': {e}")
        return None


def check_recent_reminder(loan_id, days=14):
    """
    Check if a reminder has been sent for this loan in the past 'days' days.
    """
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cutoff_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
            cursor.execute('''
            SELECT 1 FROM reminders 
            WHERE loan_id = %s AND sent_at >= %s
            ''', (loan_id, cutoff_date))
            result = cursor.fetchone()
            return result is not None


def record_reminder(loan_id):
    """
    Record that a reminder has been sent for a specific loan.
    """
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            sent_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            try:
                cursor.execute('''
                INSERT INTO reminders (loan_id, sent_at) 
                VALUES (%s, %s)
                ''', (loan_id, sent_at))
                conn.commit()
            except Exception as e:
                conn.rollback()
                print(f"Failed to insert reminder record for loan_id {loan_id}: {e}")


def fetch_last_reminder_date(loan_id):
    """
    Get the most recent reminder date for a specific loan_id.
    """
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute('''
            SELECT sent_at 
            FROM reminders 
            WHERE loan_id = %s 
            ORDER BY sent_at DESC 
            LIMIT 1
            ''', (loan_id,))
            result = cursor.fetchone()
            return result["sent_at"] if result else None


def delete_qr_code(qr_code_path):
    """
    Deletes a QR code file if the book insert fails.
    No database changes needed for this function.
    """
    try:
        if qr_code_path and os.path.exists(qr_code_path):
            os.remove(qr_code_path)
            print(f"Deleted orphaned QR code: {qr_code_path}")
    except Exception as e:
        print(f"Failed to delete orphaned QR code {qr_code_path}: {str(e)}")


def get_member_loans(member_id):
    """Retrieve books currently borrowed by a member."""
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute('''
                SELECT books.title, loans.borrowed_at
                FROM loans
                JOIN books ON loans.book_id = books.id
                WHERE loans.member_id = %s AND loans.returned_at IS NULL
            ''', (member_id,))
            loans = cursor.fetchall()
            return [{"book_title": loan[0], "borrowed_at": loan[1]} for loan in loans]


def get_member_borrowed_books_count(member_id):
    """Get the count of books currently borrowed by a member."""
    with get_postgres_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cursor:
            cursor.execute('''
                SELECT COUNT(*) AS count FROM loans
                WHERE member_id = %s AND returned_at IS NULL
            ''', (member_id,))
            result = cursor.fetchone()

            # Ensure we return a valid integer (default to 0 if no row exists)
            return result["count"] if result and "count" in result else 0


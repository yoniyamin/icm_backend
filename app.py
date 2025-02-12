# app.py
import os
from io import BytesIO
import sqlite3
from reportsServices import generate_qr_pdf, QR_CODES_DIR, extract_qr_number

import gunicorn
import logging
from datetime import datetime, timedelta, timezone
from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
import models
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
from services import database_service as db
from flask_bcrypt import Bcrypt
from functools import wraps

app = Flask(__name__)
CORS(app)
bcrypt = Bcrypt(app)

ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD_HASH = bcrypt.generate_password_hash(os.getenv("ADMIN_PASSWORD", "admin123")).decode('utf-8')
print(f"ADMIN_USERNAME: {ADMIN_USERNAME}, ADMIN_PASSWORD_HASH: {ADMIN_PASSWORD_HASH}")
TOKEN_EXPIRY_SECONDS = 3600  # 2 hour

# Initialize the database
with app.app_context():
    models.init_db()

# Simple in-memory session store (for demonstration purposes)
SESSIONS = {}


def token_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization")
        if not token or not db.validate_session_token(token):
            return jsonify({"message": "Unauthorized: Invalid or expired token."}), 403
        return f(*args, **kwargs)
    return decorated


cleanup_counter = {"count": 0}
@app.before_request
def cleanup_sessions():
    # Skip cleanup for public routes
    if request.endpoint in ["login", "get_qr_code", "get_available_books"]:
        return

    # Periodic cleanup based on a counter
    cleanup_counter["count"] += 1
    if cleanup_counter["count"] >= 100:
        print("Running periodic session cleanup")
        db.remove_expired_tokens()
        cleanup_counter["count"] = 0

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    username = data.get("username")
    password = data.get("password")

    if username == ADMIN_USERNAME and bcrypt.check_password_hash(ADMIN_PASSWORD_HASH, password):
        token = os.urandom(24).hex()
        expiry_time = datetime.now(timezone.utc) + timedelta(seconds=TOKEN_EXPIRY_SECONDS)
        expiry_time_iso = expiry_time.isoformat()  # Convert to ISO format

        print(f"Generated token: {token}, Expiry time: {expiry_time_iso}")

        # Store token in the database
        message = db.store_session_token(token, expiry_time_iso)
        print(f"Store message: {message}")

        return jsonify({"message": "Login successful", "token": token}), 200
    else:
        return jsonify({"message": "Invalid username or password"}), 401



def send_email(to_email, subject, loan_details, language='en'):
    """
    Send an email using SendGrid API with language-specific template.
    """
    template = get_reminder_template(language)
    try:
        if 'borrowed_at' in loan_details:
            loan_details['borrowed_at'] = datetime.strptime(loan_details['borrowed_at'],
                '%Y-%m-%d %H:%M:%S.%f').strftime('%Y-%m-%d')
        body = template.format(**loan_details)
    except Exception as e:
        print(f"Template formatting failed: {e}. Loan details: {loan_details}")
        return False

    # Apply RTL styles for Hebrew emails
    if language == 'he':
        body = f'<div dir="rtl" style="text-align: right; font-family: Arial, sans-serif;">{body}</div>'
    else:
        body = f'<div dir="ltr" style="text-align: left; font-family: Arial, sans-serif;">{body}</div>'

    message = Mail(
        from_email='icm.library.reminder@gmail.com',
        to_emails=to_email,
        subject=subject,
        html_content=body  # Ensure SendGrid renders it as HTML
    )
    try:
        sg = SendGridAPIClient(os.getenv('SENDGRID_API_KEY'))
        response = sg.send(message)
        print(f"Email sent successfully to {to_email} with status {response.status_code}")
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False




@app.route('/api/qr_codes/<filename>', methods=['GET'])
def get_qr_code(filename):
    """
    Serve the QR code image file.
    """
    qr_code_path = os.path.join(app.root_path, 'qr_codes', filename)
    if os.path.exists(qr_code_path):
        return send_file(qr_code_path, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404

# Route to get all books
@app.route("/api/books", methods=["GET"])
def get_books():
    # Extract 'order_by' parameter from the request
    order_by = request.args.get('order_by', 'desc')  # Default to 'desc' if not provided
    print(f"Received order_by parameter: {order_by}")  # Optional: For debugging

    # Pass 'order_by' to the database function
    books = db.get_books(order_by=order_by)
    return jsonify(books)


# Route to add a new book with all relevant fields
@app.route("/api/books", methods=["POST"])
@token_required
def add_book():
    data = request.get_json()
    title = data.get("title")
    author = data.get("author")
    description = data.get("description")
    year_of_publication = data.get("year_of_publication")
    cover_type = data.get("cover_type")
    pages = data.get("pages")
    recommended_age = data.get("recommended_age")
    book_condition = data.get("book_condition") or "טוב - בלאי קל"
    loan_status = data.get("loan_status", "available")
    delivering_parent = data.get("delivering_parent")

    print(f"Received data: {data}")
    qr_code = None

    try:
        # Call the database service to add the book (which includes generating QR code)
        qr_code = db.add_book(
            title=title,
            author=author,
            description=description,
            year_of_publication=year_of_publication,
            cover_type=cover_type,
            pages=pages,
            recommended_age=recommended_age,
            loan_status=loan_status,
            book_condition=book_condition,
            delivering_parent=delivering_parent
        )

        return jsonify({"message": "Book added successfully", "qr_code": qr_code})

    except sqlite3.IntegrityError as e:
        print(f"Database error: {str(e)}")
        if qr_code:  # ✅ Only delete if QR code was actually created
            db.delete_qr_code(qr_code)
        return jsonify({"error": f"Database error: {str(e)}"}), 400

    except Exception as e:
        print(f"Error adding book: {str(e)}")
        if qr_code:
            db.delete_qr_code(qr_code)  # Ensure the QR code is deleted
        return jsonify({"error": f"Server error: {str(e)}"}), 500


@app.route("/api/books/<int:book_id>", methods=["PUT"])
@token_required
def update_book(book_id):
    try:
        data = request.get_json()
        print(f"Updating book {book_id} with data: {data}")

        # Build update parameters
        update_fields = {}
        allowed_fields = [
            'title', 'author', 'description', 'year_of_publication',
            'cover_type', 'pages', 'recommended_age', 'book_condition',
            'delivering_parent'
        ]

        for field in allowed_fields:
            if field in data:
                update_fields[field] = data[field]

        if not update_fields:
            return jsonify({"error": "No valid fields provided for update"}), 400

        # Update the book in the database
        updated_book = db.update_book(book_id, **update_fields)

        if not updated_book:
            return jsonify({"error": "Book not found"}), 404

        return jsonify(updated_book), 200

    except sqlite3.IntegrityError as e:
        return jsonify({"error": f"Database error: {str(e)}"}), 400
    except Exception as e:
        print(f"Error updating book: {str(e)}")
        return jsonify({"error": f"Server error: {str(e)}"}), 500

# Route to update book status
@app.route("/api/books/<qr_code>/status", methods=["PUT"])
def update_book_status(qr_code):
    data = request.get_json()
    status = data.get("status")
    db.update_book_status(qr_code, status)
    return jsonify({"message": "Book status updated successfully"})


# Route to get loan history for a book
@app.route("/api/books/<int:book_id>/loans", methods=["GET"])
def get_book_loans(book_id):
    loans = db.get_book_loans(book_id)
    return jsonify(loans)

@app.route('/api/borrowing_history', methods=['GET'])
def borrowing_history():
    qr_code = request.args.get('qr_code')
    print(f"looking for {qr_code} history")
    history = db.get_borrowing_history(qr_code)
    return jsonify(history)


# Endpoint to get all members
@app.route("/api/members", methods=["GET"])
@token_required
def get_members():
    members = db.get_members()
    return jsonify(members)


# Endpoint to add a new member
@app.route("/api/members", methods=["POST"])
@token_required
def add_member():
    data = request.get_json()
    parent_name = data.get("parent_name")
    kid_name = data.get("kid_name")
    email = data.get("email")

    # Call the database service to add the member
    db.add_member(parent_name, kid_name, email)
    return jsonify({"message": "Member added successfully"}), 201


@app.route("/api/members/<int:member_id>", methods=["PUT", "DELETE"])
@token_required
def handle_member(member_id):
    if request.method == "PUT":
        data = request.get_json()
        parent_name = data.get("parent_name")
        kid_name = data.get("kid_name")
        email = data.get("email")

        db.update_member(member_id, parent_name, kid_name, email)
        return jsonify({"message": "Member updated successfully"}), 200

    elif request.method == "DELETE":
        try:
            db.delete_member(member_id)
            return jsonify({"message": "Member deleted successfully"}), 200
        except Exception as e:
            # Log the error if needed, then return an error response
            return jsonify({"error": str(e)}), 400

@app.route('/api/book/<qr_code>', methods=['GET'])
def get_book_by_qr_code(qr_code):
    book = db.get_book_by_qr_code(qr_code)
    if book:
        return jsonify(book), 200  # Explicitly set 200 for found book
    else:
        return jsonify({"error": "Book not found"}), 404  # 404 only when not found


@app.route('/api/available_books', methods=['GET'])
def get_available_books():
    available_books = db.get_books_by_status("available")
    return jsonify(available_books)


@app.route('/api/borrowed_books', methods=['GET'])
def get_borrowed_books():
    available_books = db.get_books_by_status("borrowed")
    return jsonify(available_books)


@app.route('/api/book/borrow', methods=['POST'])
@token_required
def borrow_book():
    data = request.json
    qr_code = data.get('qr_code')
    member_id = data.get('member_id')
    book_state = data.get('book_state', 'טוב - בלאי קל')  # Optional with default state

    if not qr_code or not member_id:
        return jsonify({"error": "Both qr_code and member_id are required"}), 400
    print(f"going to borrow the following: {qr_code}, {member_id}, {datetime.now()}, {book_state}")
    # Call the database service function to create a loan
    result = db.borrow_book(qr_code, member_id, datetime.now(), book_state)
    if result:
        return jsonify({"message": "Loan added successfully"})
    else:
        return jsonify({"error": "Failed to add loan"}), 500


@app.route('/api/book/return', methods=['POST'])
@token_required
def return_book():
    data = request.json
    print(data)
    qr_code = data.get('qr_code')
    result = db.return_book(qr_code)
    return jsonify(result)

@app.route('/api/open_loans', methods=['GET'])
def get_open_loans():
    qr_code = request.args.get('qr_code', None)
    loans = db.get_open_loans(qr_code=qr_code)  # Pass `None` if no QR code is provided.
    return jsonify(loans)


@app.route('/api/loans/history', methods=['GET'])
def get_loan_history():
    qr_code = request.args.get('qr_code', None)  # Get QR code, or None if not provided
    show_all = request.args.get('show_all', 'false').lower() == 'true'  # Parse show_all
    print(f"requesting for history on {qr_code}, showall is set to {show_all}")
    if qr_code:
        print(f"fetching history for {qr_code}")
        # Fetch history for a specific book
        history = db.get_loan_history(qr_code, show_all)
    else:
        # Fetch all open loans or all loans
        history = db.get_all_open_loans() if not show_all \
            else db.get_all_loans()
    print(history)
    return jsonify(history)

# Flask routes for generating reports
@app.route('/api/generate_books_report', methods=['GET'])
@token_required
def books_report():
    language = request.headers.get('Accept-Language', 'en')
    order_by = request.args.get('order_by', 'desc')
    sort_column = request.args.get('sort_column', 'title')
    include_history = request.args.get('include_history', 'true').lower() == 'true'

    report_filename = db.generate_books_report(order_by, sort_column, include_history, language)
    return send_file(report_filename, as_attachment=True)

@app.route('/api/generate_inventory_report', methods=['GET'])
@token_required
def inventory_report():
    language = request.headers.get('Accept-Language', 'en')
    order_by = request.args.get('order_by', 'desc')
    sort_column = request.args.get('sort_column', 'title')
    include_borrowed = request.args.get('include_borrowed', '').strip().lower() == 'true'

    report_filename = db.generate_inventory_report(order_by, sort_column, include_borrowed, language)
    return send_file(report_filename, as_attachment=True)


@app.route('/api/send-reminder', methods=['POST'])
@token_required
def send_reminder():
    try:
        data = request.json
        print(f"Request data: {data}")  # 🔥 Log entire request payload
        loan_id = data.get('loan_id')
        language = data.get('language', 'en')
        if not loan_id:
            print("Missing loan_id")  # 🔥 Debug
            return jsonify({"success": False, "error": "Loan ID is required"}), 400

        # 🔥 Check if a reminder was sent for this loan in the last 14 days
        if db.check_recent_reminder(loan_id, days=14):
            print(f"Reminder already sent for loan_id {loan_id} in the past 14 days.")
            return jsonify({"success": False, "error": "Reminder already sent recently"}), 400

        loan_details = data.get('loan_details', {})
        print(f"Loan details: {loan_details}")  # 🔥 Log loan details
        if not loan_details:
            print("Missing loan details")  # 🔥 Debug
            return jsonify({"success": False, "error": "Loan details are required"}), 400

        borrower_name = loan_details.get('borrower_name')
        book_title = loan_details.get('book_title')
        borrowed_at = loan_details.get('borrowed_at')

        if not all([borrower_name, book_title, borrowed_at]):
            print(f"Missing required fields: borrower_name={borrower_name}, book_title={book_title}, borrowed_at={borrowed_at}")  # 🔥 Debug
            return jsonify({"success": False, "error": "Missing borrower_name, book_title, or borrowed_at"}), 400

        email = db.find_email_by_borrower_name(borrower_name)
        if not email:
            return jsonify({"success": False, "error": "No email found for borrower"}), 404

        subject = data.get('subject', 'Reminder')
        success = send_email(email, subject, loan_details, language)

        if success:
            try:
                db.record_reminder(loan_id)
            except Exception as e:
                print(f"Error recording reminder for loan_id {loan_id}: {e}")
                return jsonify({"success": False, "error": "Failed to record reminder"}), 500

            return jsonify({"success": True, "message": f"Email sent to {email}"})
        else:
            return jsonify({"success": False, "error": f"Failed to send email to {email}"}), 500

    except Exception as e:
        print(f"Error sending reminder: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

    print("Reached end of function without return")
    return jsonify({"success": False, "error": "Unexpected server error"}), 500



@app.route('/api/reminders/last/<int:loan_id>', methods=['GET'])
def get_last_reminder(loan_id):
    """
    Get the last reminder date for a specific loan_id.
    """
    try:
        last_reminder = db.fetch_last_reminder_date(loan_id)
        if last_reminder:
            return jsonify({"sent_at": last_reminder})
        else:
            return jsonify({"message": "No reminders found"}), 200
    except Exception as e:
        print(f"Error fetching last reminder for loan_id {loan_id}: {e}")
        return jsonify({"error": str(e)}), 500


def get_reminder_template(language='en'):
    try:
        template_path = os.path.join(app.root_path, 'assets', 'template.txt')
        with open(template_path, 'r', encoding='utf-8') as file:
            templates = {}
            current_lang = None
            current_template = []

            for line in file:
                stripped_line = line.strip()  # Remove excess spaces/newlines
                if stripped_line in ['en:', 'he:']:
                    if current_lang:
                        templates[current_lang] = '\n'.join(current_template).strip()
                        current_template = []
                    current_lang = stripped_line.replace(':', '')
                elif current_lang:  # Only append if a language is selected
                    current_template.append(line.rstrip())  # Strip trailing spaces

            if current_lang:  # Save the last template
                templates[current_lang] = '\n'.join(current_template).strip()

            template = templates.get(language, templates['en'])

            # Add text direction markers for Hebrew
            if language == 'he':
                return f"\u202B{template}\u202C"  # RTL embedding
            return f"\u202A{template}\u202C"  # LTR embedding

    except Exception as e:
        print(f"Template error: {e}")
        return get_default_template(language)


def get_default_template(language):
    templates = {
        'en': """Dear {borrower_name},
                It's time to return the book "{book_title}" you borrowed on {borrowed_at}.
                Please return it as soon as possible so others can enjoy it too.

                Thanks,
                ICM Library Staff""",
        'he': """\u202Bשלום {borrower_name},
                הגיע הזמן להחזיר את הספר "{book_title}" שהושאל בתאריך {borrowed_at}.
                אנא החזירו בהקדם האפשרי כדי שאחרים יוכלו ליהנות ממנו.

                בתודה,
                צוות ספריית הקהילה הישראלית במדריד\u202C"""
    }
    return templates.get(language, templates['en'])

@app.route('/api/qr_codes', methods=['GET'])
@token_required  # Your token authentication decorator
def get_qr_codes():
    """
    Returns a JSON array of available QR code filenames from the fixed QR_CODES_DIR.
    This allows the frontend to know how many QR codes are available.
    """
    try:
        qr_files = [
            f for f in os.listdir(QR_CODES_DIR)
            if f.endswith(".png") and extract_qr_number(f) is not None
        ]
        # Optionally, sort by the numeric value extracted from the filename.
        qr_files.sort(key=lambda f: extract_qr_number(f))
        return jsonify(qr_files)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
@app.route('/api/reports/qr_codes', methods=['POST'])
@token_required
def qr_codes_pdf_report():
    """
    Expects JSON in the request body with:
      - start_qr: starting QR code number (integer).
      - end_qr: ending QR code number (integer).
    Returns a downloadable PDF file containing QR codes from the fixed qr_codes/ directory.
    """
    data = request.get_json()
    try:
        start_qr = int(data['start_qr'])
        end_qr = int(data['end_qr'])

        pdf_bytes = generate_qr_pdf(start_qr, end_qr)
        return send_file(
            BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="qr_codes_printable.pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == "__main__":
    app.run(debug=True)

import os
import psycopg2
import sqlite3

# Environment Variables
POSTGRES_URL = os.getenv("DATABASE_URL")
SQLITE_DB_PATH = "database.db"  # Keep SQLite for sessions

def init_postgres():
    """Initialize PostgreSQL for books, members, loans, and reminders."""
    with psycopg2.connect(POSTGRES_URL) as conn, conn.cursor() as cursor:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id SERIAL PRIMARY KEY,
                title TEXT NOT NULL,
                author TEXT,
                description TEXT,
                year_of_publication INTEGER,
                cover_type TEXT CHECK (cover_type IN ('כריכה רכה', 'כריכה קשה', 'עמודים קשיחים', 'ספר עם בטריה')),
                pages INTEGER,
                recommended_age INTEGER,
                book_condition TEXT CHECK (book_condition IN ('כמו חדש', 'מצויין - בלאי בלתי מורגש', 'טוב - בלאי קל')) DEFAULT 'טוב - בלאי קל',
                loan_status TEXT DEFAULT 'available' CHECK (loan_status IN ('available', 'borrowed')),
                delivering_parent TEXT,
                qr_code TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS members (
                id SERIAL PRIMARY KEY,
                parent_name TEXT NOT NULL,
                kid_name TEXT NOT NULL,
                email TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS loans (
                id SERIAL PRIMARY KEY,
                book_id INTEGER NOT NULL REFERENCES books(id),
                member_id INTEGER NOT NULL REFERENCES members(id),
                book_state TEXT CHECK (book_state IN ('כמו חדש', 'מצויין - בלאי בלתי מורגש', 'טוב - בלאי קל')) DEFAULT 'טוב - בלאי קל',
                borrowed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                returned_at TIMESTAMP DEFAULT NULL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reminders (
                id SERIAL PRIMARY KEY,
                loan_id INTEGER NOT NULL REFERENCES loans(id),
                sent_at TIMESTAMP NOT NULL,
                UNIQUE(loan_id, sent_at)
            )
        ''')

        conn.commit()

def sync_sequences():
    """Synchronize PostgreSQL sequences with the maximum ID values in their respective tables."""
    with psycopg2.connect(POSTGRES_URL) as conn, conn.cursor() as cursor:
        # Sync loans_id_seq
            cursor.execute("SELECT MAX(id) FROM loans")
            max_loan_id = cursor.fetchone()[0] or 0  # Default to 0 if table is empty
            cursor.execute("SELECT setval('loans_id_seq', %s)", (max_loan_id + 1,))

            # Sync books_id_seq (if needed)
            cursor.execute("SELECT MAX(id) FROM books")
            max_book_id = cursor.fetchone()[0] or 0
            cursor.execute("SELECT setval('books_id_seq', %s)", (max_book_id + 1,))

            # Sync members_id_seq (if needed)
            cursor.execute("SELECT MAX(id) FROM members")
            max_member_id = cursor.fetchone()[0] or 0
            cursor.execute("SELECT setval('members_id_seq', %s)", (max_member_id + 1,))

            # Sync reminders_id_seq (if needed)
            cursor.execute("SELECT MAX(id) FROM reminders")
            max_reminder_id = cursor.fetchone()[0] or 0
            cursor.execute("SELECT setval('reminders_id_seq', %s)", (max_reminder_id + 1,))

            conn.commit()
            print("Sequences synchronized successfully.")

def sync_book_loan_status():
    with psycopg2.connect(POSTGRES_URL) as conn, conn.cursor() as cursor:
        # Set books to 'available' if they have no active loans
        cursor.execute("""
            UPDATE books 
            SET loan_status = 'available'
            WHERE id NOT IN (
                SELECT book_id FROM loans WHERE returned_at IS NULL
            )
        """)

        # Set books to 'borrowed' only if they have an active loan
        cursor.execute("""
            UPDATE books 
            SET loan_status = 'borrowed'
            WHERE id IN (
                SELECT DISTINCT book_id FROM loans WHERE returned_at IS NULL
            )
        """)

        conn.commit()
        print("✅ Book loan statuses synchronized.")


def init_sqlite():
    """Initialize SQLite for session management."""
    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sessions (
                token TEXT PRIMARY KEY,
                expiry DATETIME NOT NULL
            )
        ''')
        conn.commit()

def init_db():
    """Initialize both PostgreSQL and SQLite databases."""
    init_postgres()
    init_sqlite()
    sync_sequences()
    sync_book_loan_status()

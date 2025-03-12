#!/usr/bin/env python3
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from services.database_service import generate_qr_code_with_logo, store_qr_code_in_db

# Get the database URL from the environment variable
POSTGRES_URL = os.getenv("DATABASE_URL")
if not POSTGRES_URL:
    raise ValueError("DATABASE_URL environment variable is not set")

def get_postgres_connection():
    """Establish a PostgreSQL connection using the provided URL."""
    return psycopg2.connect(POSTGRES_URL, cursor_factory=RealDictCursor)

def generate_and_insert_qr_codes():
    with get_postgres_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT b.qr_code, b.title
                FROM books b
                LEFT JOIN qr_codes q ON b.qr_code = q.qr_code
                WHERE q.qr_code IS NULL
            """)
            books_without_qr = cursor.fetchall()
            for book in books_without_qr:
                qr_code = book['qr_code']
                title = book['title']
                try:
                    image_data = generate_qr_code_with_logo(qr_code, title)
                    store_qr_code_in_db(qr_code, image_data)
                    print(f"Successfully inserted QR code for {qr_code}")
                except Exception as e:
                    print(f"Failed to generate or insert QR code for {qr_code}: {e}")

if __name__ == "__main__":
    generate_and_insert_qr_codes()
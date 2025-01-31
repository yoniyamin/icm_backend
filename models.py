import sqlite3

def init_db():
    with sqlite3.connect("database.db") as conn:
        cursor = conn.cursor()

        # Create the books table with updated fields
        cursor.execute('''
                        CREATE TABLE IF NOT EXISTS books (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            title TEXT NOT NULL,  -- Remains mandatory
                            author TEXT,          -- Changed from NOT NULL
                            description TEXT,
                            year_of_publication INTEGER,
                            cover_type TEXT CHECK (cover_type IN 
                            ('כריכה רכה', 'כריכה קשה', 'עמודים קשיחים', 'ספר עם בטריה')),
                            pages INTEGER,
                            recommended_age INTEGER,
                            book_condition TEXT CHECK (book_condition IN 
                            ('כמו חדש', 'מצויין - בלאי בלתי מורגש', 'טוב - בלאי קל')) DEFAULT 'טוב - בלאי קל',
                            loan_status TEXT DEFAULT 'available' CHECK (loan_status IN ('available', 'borrowed')),
                            delivering_parent TEXT,
                            qr_code TEXT UNIQUE NOT NULL,
                            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                        )
                        ''')

        # Create the members table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_name TEXT NOT NULL,                      -- שם ההורה
            kid_name TEXT NOT NULL,                         -- שם הילד
            email TEXT UNIQUE NOT NULL,                     -- אימייל
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        ''')

        # Create the loans table with book state options
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS loans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            book_id INTEGER NOT NULL,
            member_id INTEGER NOT NULL,
            book_state TEXT CHECK (book_state IN ('כמו חדש', 'מצויין - בלאי בלתי מורגש', 'טוב - בלאי קל')) DEFAULT 'טוב - בלאי קל', -- מצב הספר
            borrowed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            returned_at DATETIME DEFAULT NULL,
            FOREIGN KEY(book_id) REFERENCES books(id),
            FOREIGN KEY(member_id) REFERENCES members(id)
        )
        ''')

        # Create the reminders table to keep the history of when a reminder was sent
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER NOT NULL,
            sent_at TEXT NOT NULL,
            UNIQUE(loan_id, sent_at)
        )
        ''')

        # Trigger to update `updated_at` timestamp for books
        cursor.execute('''
        CREATE TRIGGER IF NOT EXISTS books_update_timestamp 
        AFTER UPDATE ON books
        BEGIN
            UPDATE books SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
        END;
        ''')

        # Trigger to update `updated_at` timestamp for members
        cursor.execute('''
        CREATE TRIGGER IF NOT EXISTS members_update_timestamp
        AFTER UPDATE ON members
        BEGIN
            UPDATE members SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
        END;
        ''')
        
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            expiry DATETIME NOT NULL
        )
        ''')

        conn.commit()

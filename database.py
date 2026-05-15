import sqlite3

def init_db(db_path='presentation.db'):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS slides (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            slide_number   INTEGER,
            title          TEXT,
            content        TEXT,
            image_prompt   TEXT,
            image_keywords TEXT,
            status         TEXT DEFAULT 'pending'
        )
    ''')
    conn.commit()
    conn.close()
    print("Database initialised.")

if __name__ == "__main__":
    init_db()

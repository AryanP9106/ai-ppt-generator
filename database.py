import sqlite3

def init_db():
    conn = sqlite3.connect('presentation.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS slides (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            slide_number INTEGER,
            title TEXT,
            content TEXT,
            image_prompt TEXT,
            status TEXT DEFAULT 'pending'
        )
    ''')
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully.")
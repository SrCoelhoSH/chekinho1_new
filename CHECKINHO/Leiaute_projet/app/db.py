import os
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = os.environ.get('DB_PATH', 'app.db')

conn = sqlite3.connect(DB_PATH)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL
)''')

c.execute('''CREATE TABLE IF NOT EXISTS results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    subfolder_name TEXT,
    report TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(user_id) REFERENCES users(id)
)''')

conn.commit()
conn.close()


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


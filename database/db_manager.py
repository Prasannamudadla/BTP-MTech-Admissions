import sqlite3
DB_NAME = "mtech_offers.db"

def get_connection():
    return sqlite3.connect(DB_NAME)

def fetch_all_candidates():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM candidates")
    rows = cursor.fetchall()
    conn.close()
    return rows

def insert_candidate(data_dict):
    conn = get_connection()
    cursor = conn.cursor()
    columns = ', '.join(data_dict.keys())
    placeholders = ', '.join(['?'] * len(data_dict))
    cursor.execute(f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})',
                   tuple(data_dict.values()))
    conn.commit()
    conn.close()

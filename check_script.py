# check_database.py
import sqlite3
import pandas as pd
import os
print("Current working dir:", os.getcwd())
print("Files here:", os.listdir())


DB_NAME = "mtech_offers.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    print(f"Connected to database: {DB_NAME}")
    return conn

def list_tables():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    conn.close()
    print("Tables in DB:", [t[0] for t in tables])

def print_candidates():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM candidates", conn)
    conn.close()
    
    if df.empty:
        print("No candidates found in the database.")
    else:
        print(f"\n=== Candidates Table ({len(df)} rows) ===")
        print(df.head(10))  # show first 10 rows
        if len(df) > 10:
            print("...")  # indicate more rows exist

def print_seat_matrix():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM seat_matrix", conn)
    conn.close()
    
    if df.empty:
        print("No seat matrix data found in the database.")
    else:
        print(f"\n=== Seat Matrix Table ({len(df)} rows) ===")
        print(df)

if __name__ == "__main__":
    list_tables()
    print_candidates()
    print_seat_matrix()


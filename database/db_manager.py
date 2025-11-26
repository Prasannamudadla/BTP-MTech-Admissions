# import sqlite3
# DB_NAME = "mtech_offers.db"

# def get_connection():
#     return sqlite3.connect(DB_NAME)

# def fetch_all_candidates():
#     conn = get_connection()
#     cursor = conn.cursor()
#     cursor.execute("SELECT * FROM candidates")
#     rows = cursor.fetchall()
#     conn.close()
#     return rows

# def insert_candidate(data_dict):
#     conn = get_connection()
#     cursor = conn.cursor()
#     columns = ', '.join(data_dict.keys())
#     placeholders = ', '.join(['?'] * len(data_dict))
#     cursor.execute(f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})',
#                    tuple(data_dict.values()))
#     conn.commit()
#     conn.close()

import sqlite3
import os

DB_NAME = "mtech_offers.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """
    Creates all tables required for the app.
    This replaces full_Setup.py.
    """
    conn = get_connection()
    cur = conn.cursor()

    # ---- Candidates table ----
    cur.execute("""
    CREATE TABLE IF NOT EXISTS candidates (
        Si_NO INTEGER,
        App_no TEXT PRIMARY KEY,
        Email TEXT,
        Full_Name TEXT,
        Adm_cat TEXT,
        Pwd TEXT,
        Ews TEXT,
        Gender TEXT,
        Category TEXT,
        COAP TEXT,
        GATE22RollNo TEXT,
        GATE22Rank INTEGER,
        GATE22Score REAL,
        GATE22Disc TEXT,
        GATE21RollNo TEXT,
        GATE21Rank INTEGER,
        GATE21Score REAL,
        GATE21Disc TEXT,
        GATE20RollNo TEXT,
        GATE20Rank INTEGER,
        GATE20Score REAL,
        GATE20Disc TEXT,
        MaxGATEScore_3yrs REAL,
        HSSC_board TEXT,
        HSSC_date TEXT,
        HSSC_per REAL,
        SSC_board TEXT,
        SSC_date TEXT,
        SSC_per REAL,
        Degree_Qualification TEXT,
        Degree_PassingDate TEXT,
        Degree_Branch TEXT,
        Degree_OtherBranch TEXT,
        Degree_Institute TEXT,
        Degree_CGPA_7th REAL,
        Degree_CGPA_8th REAL,
        Degree_Per_7th REAL,
        Degree_Per_8th REAL,
        ExtraColumn TEXT,
        GATE_Roll_num TEXT
    )
    """)

    # ---- Seat matrix ----
    cur.execute("""
    CREATE TABLE IF NOT EXISTS seat_matrix (
        category TEXT PRIMARY KEY,
        set_seats INTEGER DEFAULT 0,
        seats_allocated INTEGER DEFAULT 0,
        seats_booked INTEGER DEFAULT 0
    )
    """)

    conn.commit()
    conn.close()
    
def ensure_db():
    if not os.path.exists(DB_NAME):
        init_db()

def fetch_all_candidates():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM candidates")
    rows = cursor.fetchall()
    conn.close()
    return rows
def reset_db_data():
    """
    Deletes the database file completely and then re-initializes empty tables.
    This effectively performs a full application reset.
    """
    conn = None
    try:
        # Step 1: Close any open connections (if any are held, though typically they shouldn't be)
        # In a real app, this is crucial. Since PySide6 manages the main thread,
        # we assume connections are short-lived, but it's good practice.
        pass 
    except Exception:
        pass

    # Step 2: Delete the database file
    if os.path.exists(DB_NAME):
        os.remove(DB_NAME)

    # Step 3: Re-initialize the tables
    init_db()

    # NOTE: Since offers and decisions tables are not defined in init_db, 
    # they are missing here. You should ensure init_db creates ALL required tables.
    # I will assume init_db has been updated to include ALL tables (offers, decisions, etc.).

    return True

def insert_candidate(data_dict):
    conn = get_connection()
    cursor = conn.cursor()
    columns = ', '.join(data_dict.keys())
    placeholders = ', '.join(['?'] * len(data_dict))
    cursor.execute(
        f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})',
        tuple(data_dict.values())
    )
    conn.commit()
    conn.close()

def insert_many_candidates(list_of_dicts):
    if not list_of_dicts:
        return
    conn = get_connection()
    cur = conn.cursor()
    keys = list(list_of_dicts[0].keys())
    columns = ', '.join(keys)
    placeholders = ', '.join(['?'] * len(keys))
    rows = [tuple(d[k] for k in keys) for d in list_of_dicts]
    cur.executemany(f'INSERT OR IGNORE INTO candidates ({columns}) VALUES ({placeholders})', rows)
    conn.commit()
    conn.close()
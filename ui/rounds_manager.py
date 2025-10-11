import sqlite3
import pandas as pd
from PySide6.QtWidgets import QMessageBox

DB_NAME = "mtech_offers.db"

def run_round_1():
    """Perform Round 1 seat allocation based on MaxGATEScore_3yrs and seat availability."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    try:
        # Fetch candidates sorted by MaxGATEScore_3yrs (highest first)
        cursor.execute("""
            SELECT App_no, Full_Name, Category, Ews, Gender, Pwd, MaxGATEScore_3yrs
            FROM candidates
            WHERE MaxGATEScore_3yrs IS NOT NULL
            ORDER BY MaxGATEScore_3yrs DESC
        """)
        candidates = cursor.fetchall()
        print(f"Total candidates fetched: {len(candidates)}")  # DEBUG

        # Fetch seat matrix with exact casing
        cursor.execute("SELECT category, set_seats, seats_allocated FROM seat_matrix")
        seat_matrix = {
            cat.strip(): {"total": total or 0, "allocated": allocated or 0}
            for cat, total, allocated in cursor.fetchall()
        }
        print("Seat Matrix Loaded:", seat_matrix)  # DEBUG

        # Create offers table if not exists
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS offers (
                round_no INTEGER,
                App_no TEXT,
                Full_Name TEXT,
                category TEXT,
                MaxGATEScore_3yrs REAL,
                offer_status TEXT,
                PRIMARY KEY (round_no, App_no)
            )
        """)

        offers_made = []

        for app_no, name, base_cat, ews, gender, pwd, score in candidates:
            base_cat = base_cat.strip() if base_cat else "GEN"
            gender = gender.strip().capitalize() if gender else "Male"
            ews = ews.strip().capitalize() if ews else "No"
            pwd = pwd.strip().capitalize() if pwd else "No"

            seat_key_parts = []
            if ews == "Yes":
                seat_key_parts.append("EWS")
            else:
                seat_key_parts.append(base_cat)

            if gender == "Female":
                seat_key_parts.append("Female")
            else:
                seat_key_parts.append("FandM")

            if pwd == "Yes":
                seat_key_parts[-1] += "_PWD"

            seat_key = "_".join(seat_key_parts)

            # DEBUG: show candidate and seat_key mapping
            print(f"Candidate: {name} (App_no: {app_no}) -> Seat Key: {seat_key}")

            # Check if seat available in seat matrix
            if seat_key in seat_matrix and seat_matrix[seat_key]["allocated"] < seat_matrix[seat_key]["total"]:
                seat_matrix[seat_key]["allocated"] += 1
                offers_made.append((1, app_no, name, seat_key, score, "Offered"))
                print(f"Allocating {name} to {seat_key} (Allocated: {seat_matrix[seat_key]['allocated']}/{seat_matrix[seat_key]['total']})")  # DEBUG
            # Optional fallback for COMMON_PWD
            elif "COMMON_PWD" in seat_matrix and pwd == "Yes" and seat_matrix["COMMON_PWD"]["allocated"] < seat_matrix["COMMON_PWD"]["total"]:
                seat_matrix["COMMON_PWD"]["allocated"] += 1
                offers_made.append((1, app_no, name, "COMMON_PWD", score, "Offered"))
                print(f"Allocating {name} to COMMON_PWD (Allocated: {seat_matrix['COMMON_PWD']['allocated']}/{seat_matrix['COMMON_PWD']['total']})")  # DEBUG
            else:
                print(f"No seat available for {name} in {seat_key}")  # DEBUG

        # Insert offers into DB
        cursor.executemany("""
            INSERT OR IGNORE INTO offers (round_no, App_no, Full_Name, category, MaxGATEScore_3yrs, offer_status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, offers_made)

        # Update seat_matrix allocations in DB
        for cat, data in seat_matrix.items():
            cursor.execute("UPDATE seat_matrix SET seats_allocated = ? WHERE category = ?", (data["allocated"], cat))
        print("Seat matrix updated in DB")  # DEBUG

        conn.commit()
        QMessageBox.information(None, "Success", f"Round 1 allocation complete!\nTotal offers: {len(offers_made)}")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Error during round 1 allocation:\n{e}")
    finally:
        conn.close()


def download_offers(round_no=1):
    """Export offers for a given round to Excel."""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(f"SELECT * FROM offers WHERE round_no = {round_no}", conn)
    conn.close()

    if df.empty:
        QMessageBox.warning(None, "No Offers", f"No offers found for Round {round_no}")
        return

    filename = f"Round{round_no}_Offers.xlsx"
    df.to_excel(filename, index=False)
    QMessageBox.information(None, "Download Complete", f"Offers saved as {filename}")